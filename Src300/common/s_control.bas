Attribute VB_Name = "s_control"
Option Explicit

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


''汎用コード取得

'ユーザ定義型(def_結晶仕様3 より)
' 汎用ｺｰﾄﾞﾏｽﾀ
Public Type typ_GPCodeMaster
    'HINBAN As String * 8        ' 品番
    'MNOREVNO As Integer         ' 製品番号改訂番号
    'FACTORY As String * 1       ' 工場
    'OPECOND As String * 1       ' 操業条件
    codeNo As String * 12       ' コードＮＯ
    CODE As String * 5          ' コード
    codeCont As String          ' コード内容
    INDORDER As Long            ' 表示順
    codename As String          ' コード名称
    KUBUN As String             ' 区分
    READTIME As Double          ' リードタイム
    'IFKBN As String * 4         ' Ｉ／Ｆ区分
    'SYORIKBN As String * 1      ' 処理区分
    'SPECRRNO As String * 9      ' 仕様登録依頼番号
    'SXLMCNO As String * 12      ' ＳＸＬ製作条件番号
    'WFMCNO As String * 12       ' ＷＦ製作条件番号
    'STAFFID As String * 8       ' 社員ID
    'REGDATE As Date             ' 登録日付
    'UPDDATE As Date             ' 更新日付
    'SENDFLAG As String * 1      ' 送信フラグ
    'SENDDATE As Date            ' 送信日付
End Type

Const ERR_INVALID_MSGID = "メッセージが未登録です"




'概要      :汎用コードマスタから、特定のコードの内容文字列を引く
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :CODENO        ,I  ,String           ,コードNO
'          :CODE          ,I  ,String           ,コード
'          :戻り値        ,O  ,String           ,内容文字列
'説明      :見つからない場合はVbNullStringを返す
'履歴      :2001/06/07 作成  野村
Public Function GetGPCodeCont(ByVal codeNo As String, ByVal CODE As String) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''汎用コードマスタから、特定のコードの内容文字列を引く
    sql = "select CODECONT from TBCME033 where (rtrim(CODENO)='" & Trim$(codeNo) & "') and (rtrim(CODE)='" & Trim$(CODE) & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、VbNullStringを返す
        GetGPCodeCont = vbNullString
    Else
        ''見つかったら、コード内容文字列を返す
        GetGPCodeCont = rs("CODECONT")
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
End Function


'概要      :汎用コードマスタから、特定のコードを引く
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :CODENO        ,I  ,String           ,コードNO
'          :CODE          ,I  ,String           ,コード
'          :GPCode        ,O  ,typ_GPCodeMaster ,対応するデータ
'          :戻り値        ,O  ,FUNCTION_RETURN  ,成功/失敗
'説明      :
'履歴      :2001/06/04 作成  野村
Public Function GetGPCode(ByVal codeNo As String, ByVal CODE As String, GPCode As typ_GPCodeMaster) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''汎用コードマスタから、特定のコードを引く
    sql = "select CODECONT, CODENAME, INDORDER, KUBUN, READTIME from TBCME033 " & _
          "where (rtrim(CODENO)='" & Trim$(codeNo) & "') and (rtrim(CODE)='" & Trim$(CODE) & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、データ内容を消去してFUNCTION_RETURN_FAILUREを返す
        With GPCode
            .codeNo = vbNullString
            .CODE = vbNullString
            .codeCont = vbNullString
            .codename = vbNullString
            .INDORDER = 0
            .KUBUN = vbNullString
            .READTIME = 0
        End With
        GetGPCode = FUNCTION_RETURN_FAILURE
    Else
        ''見つからなかったら、データ内容を設定してFUNCTION_RETURN_SUCCESSを返す
        With GPCode
            .codeNo = codeNo
            .CODE = CODE
            .codeCont = rs("CODECONT")
            .codename = rs("CODENAME")
            .INDORDER = rs("INDORDER")
            .KUBUN = rs("KUBUN")
            .READTIME = rs("READTIME")
        End With
        GetGPCode = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
End Function


'概要      :汎用コードマスタから、コードNOに対応するコードの一覧を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :CODENO        ,I  ,String           ,コードNO
'          :GPCodeList()  ,O  ,typ_GPCodeMaster ,対応するコードデータの一覧
'          :戻り値        ,O  ,Integer          ,成功/失敗
'説明      :
'履歴      :2001/06/04 作成  野村
Public Function GetGPCodeList(ByVal codeNo As String, GPCodeList() As typ_GPCodeMaster) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim i As Integer
Dim recCnt As Integer

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''汎用コードマスタから、コードNOに対応するコードの一覧を得る
    sql = "select CODE, CODECONT, CODENAME, INDORDER, KUBUN, READTIME from TBCME033 where (rtrim(CODENO)='" & Trim$(codeNo) & "') order by INDORDER"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.EOF Then
        ''見つからなかったら、0件としてFUNCTION_RETURN_FAILUREを返す
        ReDim GPCodeList(0)
        GetGPCodeList = FUNCTION_RETURN_FAILURE
    Else
        ''見つかったら、その件数分のデータをコピーしてFUNCTION_RETURN_SUCCESSを返す
        recCnt = rs.RecordCount
        ReDim GPCodeList(recCnt)
        For i = 1 To recCnt
            With GPCodeList(i)
                .codeNo = codeNo
                .CODE = rs("CODE")
                .codeCont = rs("CODECONT")
                .codename = rs("CODENAME")
                .INDORDER = rs("INDORDER")
                .KUBUN = rs("KUBUN")
                .READTIME = rs("READTIME")
                rs.MoveNext
            End With
        Next
        GetGPCodeList = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
End Function


'概要      :コンボ等に表示する「コード:コード内容」の文字列を生成する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :code          ,I  ,String    ,コード
'          :codeCont      ,I  ,String    ,コード内容
'          :戻り値        ,O  ,String    ,結果文字列
'説明      :
'履歴      :2001/06/07 作成  野村
Public Function GetGPCodeDspStr(CODE$, codeCont$) As String

    If (Trim$(CODE) = "SPACE") Or (Trim$(CODE) = vbNullString) Then
        ''コードが「SPACE」の場合、結果文字列=" "とする
        GetGPCodeDspStr = " "
    Else
        ''それ以外の場合、文字列をつなぎ合わせる
        GetGPCodeDspStr = Trim$(CODE) & ":" & Trim$(codeCont)
    End If
End Function


'概要      :コンボボックスに汎用マスタ内の選択肢を設定する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :GPCodeList()  ,I  ,typ_GPCodeMaster ,選択肢のリスト
'          :cmb           ,I  ,ComboBox         ,設定先のコンボボックス
'          :戻り値        ,O  ,FUNCTION_RETURN  ,成功/失敗
'説明      :
'履歴      :2001/06/04 作成  野村
Public Function SetGPCodeList2Combo(GPCodeList() As typ_GPCodeMaster, cmb As ComboBox) As FUNCTION_RETURN
Dim RET As FUNCTION_RETURN
Dim max As Integer
Dim i As Integer

    With cmb
        .Clear
        max = UBound(GPCodeList)
        For i = 1 To max
            .AddItem GetGPCodeDspStr(GPCodeList(i).CODE, GPCodeList(i).codeCont)
        Next
    End With
    SetGPCodeList2Combo = FUNCTION_RETURN_SUCCESS
End Function


'概要      :コンボボックスに汎用マスタ内の選択肢を設定する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :CODENO        ,I  ,String           ,コードNO
'          :cmb           ,I  ,ComboBox         ,設定先のコンボボックス
'          :戻り値        ,O  ,FUNCTION_RETURN  ,成功/失敗
'説明      :
'履歴      :2001/06/28 作成  野村
Public Function SetGPCode2Combo(ByVal codeNo As String, cmb As ComboBox) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim i As Integer
Dim recCnt As Integer

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''汎用コードマスタから、コードNOに対応するコードの一覧を得る
    sql = "select CODE, CODECONT from TBCME033 where (rtrim(CODENO)='" & Trim$(codeNo) & "') order by INDORDER"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.EOF Then
        ''見つからなかったら、0件としてFUNCTION_RETURN_FAILUREを返す
        SetGPCode2Combo = FUNCTION_RETURN_FAILURE
    Else
        ''見つかったら、その件数分のデータをコピーしてFUNCTION_RETURN_SUCCESSを返す
        recCnt = rs.RecordCount
        cmb.Clear
        For i = 1 To recCnt
            cmb.AddItem GetGPCodeDspStr(rs("CODE"), rs("CODECONT"))
            rs.MoveNext
        Next
        SetGPCode2Combo = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
End Function


''結晶操業コード取得
''参照:s_cmzcTBCMB005_SQL.bas (コードマスタ)

'ユーザ定義型


'概要      :結晶操業用コードを検索し、指定フィールドの内容を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SYSCLASS      ,I  ,String    ,システム区分
'          :CLASS         ,I  ,String    ,区分
'          :CODE          ,I  ,String    ,コード
'          :FieldName     ,I  ,String    ,フィールド名
'          :戻り値        ,O  ,String    ,フィールド内容
'説明      :
'履歴      :2001/06/14 作成  野村
Public Function GetCodeField(ByVal SYSCLASS$, ByVal Class$, ByVal CODE$, ByVal FieldName$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetCodeField"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''コードマスタから、指定のフィールドを引く
    sql = "SELECT " & FieldName & " from TBCMB005 WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "') and (CODE='" & CODE & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、VbNullStringを返す
        GetCodeField = vbNullString
    Else
        ''見つかったら、指定フィールドの内容を返す
        GetCodeField = Trim$(rs(FieldName))
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :結晶操業用コードを検索し、指定フィールドの内容を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SYSCLASS      ,I  ,String    ,システム区分
'          :CLASS         ,I  ,String    ,区分
'          :FieldName     ,I  ,String    ,フィールド名
'          :FieldData()   ,O  ,String    ,フィールド内容
'          :戻り値        ,O  ,String    ,フィールド内容
'説明      :
'履歴      :2001/07/26 作成  野村
Public Function GetCodeField2(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$, FieldData() As String) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim i As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetCodeField2"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''コードマスタから、指定のフィールドを引く
    sql = "SELECT " & FieldName & " from TBCMB005 WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    ReDim FieldData(rs.RecordCount)
    For i = 1 To rs.RecordCount
        FieldData(i) = rs(FieldName)
        rs.MoveNext
    Next
    GetCodeField2 = FUNCTION_RETURN_SUCCESS
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    GetCodeField2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶操業用コードを検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :SYSCLASS      ,I  ,String       ,システム区分
'          :CLASS         ,I  ,String       ,区分
'          :CODE          ,I  ,String       ,コード
'          :CodeData      ,O  ,typ_TBCMB005 ,検索結果
'          :戻り値        ,O  ,FUNCTION_RETURN,検索の成否
'説明      :
'履歴      :2001/06/07 作成  野村
Public Function GetCode(ByVal SYSCLASS$, ByVal Class$, ByVal CODE$, CodeData As typ_TBCMB005) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim sqlWhere As String
Dim rec() As typ_TBCMB005
Dim RET As FUNCTION_RETURN



    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetCode"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''DB関数を利用して、コード内容を取得する
    sqlWhere = "where (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "') and (CODE='" & CODE & "')"
    RET = DBDRV_GetTBCMB005(rec, sqlWhere)
    If (UBound(rec) = 0) Then
        GetCode = FUNCTION_RETURN_FAILURE
        CodeData = rec(0)
    Else
        GetCode = FUNCTION_RETURN_SUCCESS
        CodeData = rec(1)
    End If
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶操業用コードの一覧を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :SYSCLASS      ,I  ,String       ,システム区分
'          :CLASS         ,I  ,String       ,区分
'          :CodeList()    ,O  ,typ_TBCMB005 ,コード内容の一覧
'          :戻り値        ,O  ,FUNCTION_RETURN,検索の成否
'説明      :
'履歴      :
Public Function GetCodeList(ByVal SYSCLASS$, ByVal Class$, CodeList() As typ_TBCMB005) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim sqlWhere As String
Dim sqlOrder As String
Dim RET As FUNCTION_RETURN



    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetCodeList"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''DB関数を利用して、コード内容を取得する
    sqlWhere = " where (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')"
    sqlOrder = " Order by INFO9,CODE"
    RET = DBDRV_GetTBCMB005(CodeList, sqlWhere, sqlOrder)
    GetCodeList = RET
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function
'概要      :結晶操業用コードの一覧を得る   2006/01
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :SYSCLASS      ,I  ,String       ,システム区分
'          :CLASS         ,I  ,String       ,区分
'          :INFO          ,I  ,String       ,表示制限
'          :CodeList()    ,O  ,typ_TBCMB005 ,コード内容の一覧
'          :戻り値        ,O  ,FUNCTION_RETURN,検索の成否
'説明      :
'履歴      :
Public Function GetCodeListSC18(ByVal SYSCLASS$, ByVal Class$, ByVal INFO$, CodeList() As typ_TBCMB005) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim sqlWhere As String
Dim sqlOrder As String
Dim RET As FUNCTION_RETURN



    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetCodeListSC18"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''DB関数を利用して、コード内容を取得する
    If INFO = "CM" Then
        sqlWhere = " where (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')and (trim(INFO2) ='" & INFO & "')"
    ElseIf INFO = "外形" Then
        '08/12/23 ooba
        sqlWhere = " where (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')and (trim(INFO4) ='" & INFO & "')"
    Else
        sqlWhere = " where (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')and (trim(INFO3) ='" & INFO & "')"
    End If
    sqlOrder = " Order by INFO9,CODE"
    RET = DBDRV_GetTBCMB005(CodeList, sqlWhere, sqlOrder)
    GetCodeListSC18 = RET
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :社員IDから社員氏名を求める(TBCMB001より取得)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :staffID       ,I  ,String    ,社員ID
'          :戻り値        ,O  ,String    ,社員氏名
'説明      :見つからなかった場合は、VbNullStringを返す
'履歴      :2001/06/07 作成  野村
Public Function GetStaffName(StaffID$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetStaffName"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''社員マスタから、社員名を引く
    sql = "SELECT JFMLNAME, JFSTNAME from TBCMB001 WHERE (STAFFID='" & StaffID & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、VbNullStringを返す
        GetStaffName = vbNullString
    Else
        ''見つかったら、氏名を返す
        GetStaffName = Trim$(rs("JFMLNAME")) & " " & Trim$(rs("JFSTNAME"))
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :社員IDから社員氏名を求める(テーブルKODA9より取得)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :staffID       ,I  ,String    ,社員ID
'          :戻り値        ,O  ,String    ,社員氏名
'説明      :見つからなかった場合は、VbNullStringを返す
'           2009/09/04 SUMCO Akizuki
'                      CMBC052を参考に作成

Public Function GetStaffName_KODA9(StaffID$) As String
    Dim dbIsMine As Boolean
    Dim rs As OraDynaset
    Dim sql As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cntrol.bas -- Function newGetStaffName"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''社員マスタから、社員名を引く
    sql = ""
    sql = "select NAMEJA9 from KODA9 "
    sql = sql & " where SYSCA9='K' and SHUCA9='55' and CODEA9='" & StaffID & "'"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        ''見つからなかったら、VbNullStringを返す
        GetStaffName_KODA9 = vbNullString
    Else
        ''見つかったら、氏名を返す
        GetStaffName_KODA9 = Trim$(rs("NAMEJA9"))
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :社員IDから社員氏名を求める(200mm接続時、cmcc100、cmec053用)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :staffID       ,I  ,String    ,社員ID
'          :戻り値        ,O  ,String    ,社員氏名
'説明      :見つからなかった場合は、VbNullStringを返す
'履歴      :2008/07/07　SET 小柴 作成
Public Function GetStaffName200(StaffID$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetStaffName200"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''社員マスタから、社員名を引く
    sql = "SELECT NAMEJA9 from KODA9 WHERE CODEA9='" & StaffID & _
    "' and SYSCA9='K' and SHUCA9='55'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、VbNullStringを返す
        GetStaffName200 = vbNullString
    Else
        ''見つかったら、氏名を返す
        GetStaffName200 = Trim$(rs("NAMEJA9"))
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :コンボボックスに結晶操業マスタ内の選択肢を設定する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SYSCLASS      ,I  ,String    ,SYS区分
'          :Class         ,I  ,String    ,区分
'          :FieldName     ,I  ,String    ,選択肢名称フィールド名
'          :cmb           ,O  ,ComboBox  ,設定先のコンボボックス
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :
'履歴      :2001/06/28 作成  野村
Public Function SetCode2Combo(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$, cmb As ComboBox) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function SetCode2Combo"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''コードマスタから、指定のフィールドを引く
    sql = "SELECT CODE, " & FieldName & " from TBCMB005" & _
          " WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')" & _
          " Order by INFO9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、NG
        cmb.Clear
        SetCode2Combo = FUNCTION_RETURN_FAILURE
    Else
        ''見つかったら、コンボボックスに選択肢を設定する
        With cmb
            .Clear
            max = rs.RecordCount
            For i = 1 To max
                .AddItem GetGPCodeDspStr(rs("CODE"), rs(FieldName))
                rs.MoveNext
            Next
        End With
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function
'概要      :コンボボックスに結晶操業マスタ内の選択肢を設定する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SYSCLASS      ,I  ,String    ,SYS区分
'          :Class         ,I  ,String    ,区分
'          :FieldName     ,I  ,String    ,選択肢名称フィールド名
'          :cmb           ,O  ,ComboBox  ,設定先のコンボボックス
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :
'履歴      :2001/06/28 作成  野村
Public Function SetCode2ComboSC18(ByVal SYSCLASS$, ByVal Class$, ByVal INFO$, cmb As ComboBox) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function SetCode2ComboSC18"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''コードマスタから、指定のフィールドを引く
    sql = "SELECT CODE, INFO1 from TBCMB005" & _
          " WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')"
    If INFO = "CM" Then
        sql = sql & " AND   (INFO2 ='" & INFO & "')"
    Else
        sql = sql & " AND   (INFO3 ='" & INFO & "')"
    End If
    sql = sql & " Order by INFO9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、NG
        cmb.Clear
        SetCode2ComboSC18 = FUNCTION_RETURN_FAILURE
    Else
        ''見つかったら、コンボボックスに選択肢を設定する
        With cmb
            .Clear
            max = rs.RecordCount
            For i = 1 To max
                .AddItem GetGPCodeDspStr(rs("CODE"), rs("INFO1"))
                rs.MoveNext
            Next
        End With
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :コンボボックスに結晶操業マスタ内の選択肢を設定する（":コード内容"なしバージョン）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SYSCLASS      ,I  ,String    ,SYS区分
'          :Class         ,I  ,String    ,区分
'          :FieldName     ,I  ,String    ,選択肢名称フィールド名
'          :cmb           ,O  ,ComboBox  ,設定先のコンボボックス
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :
'履歴      :2001/08/21 作成  蔵本
Public Function SetCode2Combo2(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$, cmb As ComboBox) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer
Dim CODE As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function SetCode2Combo2"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''コードマスタから、指定のフィールドを引く
    sql = "SELECT CODE, " & FieldName & " from TBCMB005" & _
          " WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')" & _
          " Order by INFO9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、NG
        cmb.Clear
        SetCode2Combo2 = FUNCTION_RETURN_FAILURE
    Else
        ''見つかったら、コンボボックスに選択肢を設定する
        With cmb
            .Clear
            max = rs.RecordCount
            For i = 1 To max
                CODE = rs("CODE")
                '.AddItem GetGPCodeDspStr(rs("CODE"), rs(FieldName))
                If (Trim$(CODE) = "SPACE") Or (Trim$(CODE) = vbNullString) Then
                     ''コードが「SPACE」の場合、結果文字列=" "とする
                    CODE = " "
                Else
                    CODE = Trim$(CODE)
                End If
                .AddItem CODE
                rs.MoveNext
            Next
        End With
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :SPREADのコンボボックスに設定するため、結晶操業マスタ内の選択肢文字列を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SYSCLASS      ,I  ,String    ,SYS区分
'          :Class         ,I  ,String    ,区分
'          :FieldName     ,I  ,String    ,選択肢名称フィールド名
'          :戻り値        ,O  ,String    ,選択肢文字列
'説明      :
'履歴      :2001/06/28 作成  野村
Public Function GetSSComboStr(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer
Dim cmbStr As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetSSComboStr"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''コードマスタから、指定のフィールドを引く
    sql = "SELECT CODE, " & FieldName & " from TBCMB005" & _
          " WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')" & _
          " Order by INFO9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、NG
        GetSSComboStr = vbNullString
    Else
        ''見つかったら、選択肢文字列を設定する
        max = rs.RecordCount
        For i = 1 To max
            If cmbStr <> vbNullString Then
                cmbStr = cmbStr & vbTab
            End If
            cmbStr = cmbStr & GetGPCodeDspStr(rs("CODE"), rs(FieldName))
            rs.MoveNext
        Next
        GetSSComboStr = cmbStr
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function
'概要      :SPREADのコンボボックスに設定するため、結晶操業マスタ内の選択肢文字列を取得する(KODA9版)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SYSCLASS      ,I  ,String    ,SYS区分
'          :Class         ,I  ,String    ,区分
'          :FieldName     ,I  ,String    ,選択肢名称フィールド名
'          :戻り値        ,O  ,String    ,選択肢文字列
'説明      :
'履歴      :2001/06/28 作成  野村
Public Function GetSSComboStrA9(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer
Dim cmbStr As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetSSComboStrA9"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''コードマスタから、指定のフィールドを引く
    sql = "SELECT CODEA9, " & FieldName & " from KODA9" & _
          " WHERE (SYSCA9='" & SYSCLASS & "') and (SHUCA9='" & Class & "')" & _
          " Order by CTR01A9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、NG
        GetSSComboStrA9 = vbNullString
    Else
        ''見つかったら、選択肢文字列を設定する
        max = rs.RecordCount
        For i = 1 To max
            If cmbStr <> vbNullString Then
                cmbStr = cmbStr & vbTab
            End If
            cmbStr = cmbStr & GetGPCodeDspStr(rs("CODEA9"), rs(FieldName))
            rs.MoveNext
        Next
        GetSSComboStrA9 = cmbStr
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :メッセージ文字列を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :MsgID         ,I  ,String    ,メッセージID
'          :params()      ,I  ,Variant   ,埋め込みパラメータ(必要な数だけ)
'          :戻り値        ,O  ,String    ,メッセージ文字列
'説明      :取得できなかった場合は、固定メッセージ(メッセージをDBにありません)を返す
'履歴      :2001/06/07 作成  野村
Public Function GetMsgStr(ByVal MsgID$, ParamArray params() As Variant) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset    'レコードセット
Dim sql As String       'SQL文字列
Dim fmt As String       '書式文字列


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004c.bas -- Function GetMsgStr"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''メッセージマスターから書式文字列を取得する
    sql = "select FORMINFO from TBCMB003 where (MSGID='" & MsgID & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''該当するメッセージIDがないときは、固定メッセージを返す
        GetMsgStr = ERR_INVALID_MSGID & "(" & MsgID & ")"
    Else
        ''通常は、書式にパラメータを埋め込んで返す
        fmt = rs("FORMINFO")
        GetMsgStr = FmtStr(fmt, params)
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :メッセージ文字列を取得する(200mm接続時、cmcc100、cmec053用)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :MsgID         ,I  ,String    ,メッセージID
'          :params()      ,I  ,Variant   ,埋め込みパラメータ(必要な数だけ)
'          :戻り値        ,O  ,String    ,メッセージ文字列
'説明      :取得できなかった場合は、固定メッセージ(メッセージをDBにありません)を返す
'          :200mmDB接続時はDBLINK経由で300mmDBのテーブルを参照する。
'履歴      :2008/07/09 作成  野村
Public Function GetMsgStr200(ByVal MsgID$, ParamArray params() As Variant) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset    'レコードセット
Dim sql As String       'SQL文字列
Dim fmt As String       '書式文字列


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004c.bas -- Function GetMsgStr"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''メッセージマスターから書式文字列を取得する
    sql = "select FORMINFO from TBCMB003@DBLINK300 where (MSGID='" & MsgID & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''該当するメッセージIDがないときは、固定メッセージを返す
        GetMsgStr200 = ERR_INVALID_MSGID & "(" & MsgID & ")"
    Else
        ''通常は、書式にパラメータを埋め込んで返す
        fmt = rs("FORMINFO")
        GetMsgStr200 = FmtStr(fmt, params)
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :指定の書式にパラメータを埋め込んだ文字列を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :fmt           ,I  ,String    ,埋め込み先文字列(printf書式風)
'          :params()      ,I  ,Variant   ,埋め込みパラメータ(可変個数)
'          :戻り値        ,O  ,String    ,埋め込み結果文字列
'説明      :
'履歴      :2001/06/06 作成  長野
Private Function FmtStr(ByVal fmt$, ParamArray params() As Variant) As String
Dim w_str       As String       '埋め込み文字列
Dim w_wrd       As String       'ﾊﾟﾗﾒｰﾀ文字列
Dim i           As Integer      'ﾙｰﾌﾟｶｳﾝﾄ
Dim n           As Integer      'ﾊﾟﾗﾒｰﾀｶｳﾝﾄ
Dim Str_Value() As String       'ﾌｫｰﾏｯﾄを｢%｣毎に区切った配列
Dim s_max       As Integer      '添字の最大値(ﾌｫｰﾏｯﾄ配列)


    '引数のﾌｫｰﾏｯﾄ文字列を｢%｣毎に区切って配列に格納

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc004c.bas -- Function FmtStr"

    Str_Value() = Split(fmt, "%")
    s_max = UBound(Str_Value)           '添字の最大値取得
    
    n = 0                               'ﾊﾟﾗﾒｰﾀ配列の添字ｶｳﾝﾄ初期化
    
    '文字列作成
    For i = 0 To s_max                  '｢%｣で区切ったﾌｫｰﾏｯﾄ文字列数分ﾙｰﾌﾟ
        '%の次の文字によって処理を分岐
        Select Case Left(Str_Value(i), 1)
        Case "s"                        '文字列の場合
            If n > UBound(params(0)) Then
                w_wrd = vbNullString
            Else
                w_wrd = params(0)(n)        'ﾊﾟﾗﾒｰﾀ文字列の取得
            End If
            w_str = w_str & w_wrd & Mid(Str_Value(i), 2)
            n = n + 1                   '次のﾊﾟﾗﾒｰﾀへ
        Case ""                         '｢%｣文字の場合
            'If i = s_max Then           '文節が最後の場合
                w_str = w_str & Str_Value(i) 'Mid(Str_Value(i), 2)
            'Else                        '文節が後に続く場合
            '    w_str = w_str & "%"     '｢%｣を代入
            '    i = i + 1               '次の節は飛ばす
            'End If
        Case Else
            w_str = w_str & Str_Value(i)
        End Select
    Next
       
    FmtStr = w_str

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :8桁品番に対応する最新の品番情報を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型          ,説明
'          :hinban        ,I  ,String      ,8桁品番
'          :fullHinban    ,O  ,tFullHinban ,品番情報
'          :[chkUsable]   ,I  ,Boolean     ,使用開始Tbl と 結晶内側管理Tbl を必須とするか
'          :戻り値        ,O  ,FUNCTION_RETURN,検索の成否
'説明      :同一の8桁品番の中では、改訂番号が大きいものがより新しい
'          :更にその中では操業条件番号が大きいものがより新しい
'履歴      :2001/06/07 作成  野村
Public Function GetLastHinban(hinban$, fullHinban As tFullHinban, Optional chkUsable As Boolean = True) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetLastHinban"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''製品仕様SXLデータ1 から、指定品番のレコードを新しい順に取り出す
    ''最も「新しい」データは、改訂番号が最大のものの中で操業条件番号が最大であるレコードである
    ''(ただし、使用開始Tblと結晶内側管理Tblに登録されていない品番は、まだ利用不能である)8/22修正
    ''2006/09 TBCME036開始フラグが'1'(開始電文を受信)のもの <---- 2006/12/11 修正
    If chkUsable Then
        'sql = "select E018.HINBAN, E018.MNOREVNO, E018.FACTORY, E018.OPECOND " & _
              "from TBCME018 E018, TBCME032 E032, TBCME036 E036 " & _
              "Where (E018.HINBAN = '" & hinban & "')" & _
              " and (E018.HINBAN=E032.HINBAN) and (E018.MNOREVNO=E032.MNOREVNO)" & _
              " and (E018.FACTORY=E032.FACTORY) " & _
              " and (E018.HINBAN=E036.HINBAN) and (E018.MNOREVNO=E036.MNOREVNO)" & _
              " and (E018.FACTORY=E036.FACTORY) and (E018.OPECOND=E036.OPECOND) " & _
              "order by MNOREVNO DESC, OPECOND DESC"
        'sql = "select E018.HINBAN, E018.MNOREVNO, E018.FACTORY, E018.OPECOND " & _
              "from TBCME018 E018 " & _
              "Where (E018.HINBAN = '" & hinban & "')" & _
              " and (E018.SYNFLAG IS NULL OR E018.SYNFLAG='1') " & _
              " and (E018.OPECOND <> '1') " & _
              "order by MNOREVNO DESC, OPECOND DESC"
        sql = "select E018.HINBAN, E018.MNOREVNO, E018.FACTORY, E018.OPECOND " & _
              "from TBCME018 E018 , TBCME036 E036 " & _
              "Where (E018.HINBAN = '" & hinban & "')" & _
              " and (E018.SYNFLAG IS NULL OR E018.SYNFLAG='1') " & _
              " and (E018.OPECOND <> '1') " & _
              " and (E018.HINBAN=E036.HINBAN) and (E018.MNOREVNO=E036.MNOREVNO)" & _
              " and (E018.FACTORY=E036.FACTORY) and (E018.OPECOND=E036.OPECOND) " & _
              " and (E036.KAISIFLG = '1') " & _
              "order by MNOREVNO DESC, OPECOND DESC"
    Else
        'こちらは仕様受入・製作条件入力用
        '製作条件付与取消に登録されている品番は無効とする
        sql = "select A.HINBAN, A.MNOREVNO, A.FACTORY, A.OPECOND " & _
              "from TBCME018 A, TBCME031 B " & _
              "where (A.HINBAN = '" & hinban & "')" & _
              " and (A.HINBAN=B.HINBAN(+)) and (A.MNOREVNO=B.MNOREVNO(+)) and (A.FACTORY=B.FACTORY(+))" & _
              " and (B.HINBAN is null) " & _
              "order by MNOREVNO DESC, OPECOND DESC"
    End If
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
    'If rs.RecordCount = 0 Or rs("OPECOND") = "1" Then
        ''見つからなかったら、FUNCTION_RETURN_FAILUREを返す
        With fullHinban
            .hinban = vbNullString
            .mnorevno = 0
            .factory = vbNullString
            .opecond = vbNullString
        End With
        GetLastHinban = FUNCTION_RETURN_FAILURE
    Else
        ''見つかったら、最新の品番情報をセットして FUNCTION_RETURN_SUCCESSを返す
        With fullHinban
            .hinban = rs("HINBAN")
            .mnorevno = rs("MNOREVNO")
            .factory = rs("FACTORY")
            .opecond = rs("OPECOND")
        End With
        GetLastHinban = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :8桁品番に対応する最新の仕様品番情報を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型          ,説明
'          :hinban        ,I  ,String      ,8桁品番
'          :fullHinban    ,O  ,tFullHinban ,品番情報
'          :戻り値        ,O  ,FUNCTION_RETURN,検索の成否
'説明      :仕様データの操業条件は常に「1」である
'          :同一の8桁品番の中では、改訂番号が大きいものがより新しい
'履歴      :2001/06/07 作成  野村
Public Function GetLastSpecHinban(hinban$, fullHinban As tFullHinban) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetLastSpecHinban"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''製品仕様SXLデータ1 から、指定品番の仕様レコードを新しい順に取り出す
    ''最も「新しい」データは、改訂番号が最大のものの中であるレコードである
    sql = "SELECT HINBAN, MNOREVNO, FACTORY, OPECOND " & _
          "From TBCME018 " & _
          "Where (HINBAN = '" & hinban & "') AND (OPECOND = '1')" & _
          "ORDER BY MNOREVNO DESC;"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、FUNCTION_RETURN_FAILUREを返す
        With fullHinban
            .hinban = vbNullString
            .mnorevno = 0
            .factory = vbNullString
            .opecond = vbNullString
        End With
        GetLastSpecHinban = FUNCTION_RETURN_FAILURE
    Else
        ''見つかったら、最新の品番情報をセットして FUNCTION_RETURN_SUCCESSを返す
        With fullHinban
            .hinban = rs("HINBAN")
            .mnorevno = rs("MNOREVNO")
            .factory = rs("FACTORY")
            .opecond = rs("OPECOND")
        End With
        GetLastSpecHinban = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :指定の結晶番号に含まれる品番の一覧を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型          ,説明
'          :cryno         ,I  ,String      ,結晶番号
'          :hinban()      ,O  ,tFullHinban ,品番リスト
'          :戻り値        ,O  ,FUNCTION_RETURN,抽出の成否
'説明      :
'履歴      :2001/06/27 作成  長野
Public Function GetXlHinban(cryno$, hinban() As tFullHinban) As FUNCTION_RETURN
Dim rs      As OraDynaset               '抽出RecordDynaset
Dim rsCnt   As Integer                  'ﾚｺｰﾄﾞｶｳﾝﾄ
Dim sql     As String                   'SQL文
Dim i       As Integer                  'ﾙｰﾌﾟｶｳﾝﾄ

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetXlHinban"

    'SQL文の作成
    sql = "Select CRYNUM, HINBAN, REVNUM, FACTORY, OPECOND from TBCME041 "
    sql = sql & "Where(CRYNUM = '" & cryno & "')"
    
    'データの抽出
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '''抽出レコードが存在しない場合
    If rs.EOF Then
        ReDim hinban(0)                     '配列の初期化
        GetXlHinban = FUNCTION_RETURN_FAILURE   'ｴﾗｰｽﾃｰﾀｽ
        GoTo proc_exit
    End If
        
    rsCnt = rs.RecordCount                  'ﾚｺｰﾄﾞ数のｶｳﾝﾄを取る
    ReDim hinban(rsCnt - 1)                 '配列の再定義
    
    '配列に値をセット
    rs.MoveFirst                            '先頭ﾚｺｰﾄﾞに移動
    For i = 0 To rsCnt - 1                  'ﾚｺｰﾄﾞ数分ﾙｰﾌﾟ
        DoEvents
        With hinban(i)
            .hinban = rs!hinban             '品番
            .mnorevno = rs!REVNUM           '製品番号改訂番号
            .factory = rs!factory           '工場
            .opecond = rs!opecond           '操業条件
        End With
        rs.MoveNext                         '次ﾚｺｰﾄﾞに移動
    Next
    
    GetXlHinban = FUNCTION_RETURN_SUCCESS   '正常ｽﾃｰﾀｽ
 

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :ドーパント濃度マスタからドーパント名の一覧を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :DopeName()    ,O  ,String    ,ドーパント名
'          :戻り値        ,O  ,FUNCTION_RETURN,抽出の成否
'説明      :
'履歴      :2001/08/08 作成  野村
'          :2011/05/09 取得ＤＢ変更 Kameda
Public Function GetDopeNames(DopeName() As String) As FUNCTION_RETURN
Dim rs      As OraDynaset               '抽出RecordDynaset
Dim rsCnt   As Integer                  'ﾚｺｰﾄﾞｶｳﾝﾄ
Dim sql     As String                   'SQL文
Dim i       As Integer                  'ﾙｰﾌﾟｶｳﾝﾄ

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetDopeNames"

    GetDopeNames = FUNCTION_RETURN_FAILURE
    
    'SQL文の作成
    'sql = "select DOPKIND from TBCMB009 order by DOPKIND"    2011/05/09 Kameda
    'SQL編集
    sql = "SELECT  NVL(codea9   , ' ') DOPKIND "
    sql = sql & "  FROM koda9 "
    sql = sql & " WHERE sysca9 = 'X'"
    sql = sql & "   AND shuca9 = 'D0'"
    sql = sql & " ORDER BY codea9"
    'データの抽出
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '''抽出レコードが存在しない場合
    If rs.EOF Then
        ReDim DopeName(0)                     '配列の初期化
        GoTo proc_exit
    End If
        
    rsCnt = rs.RecordCount                  'ﾚｺｰﾄﾞ数のｶｳﾝﾄを取る
    ReDim DopeName(1 To rsCnt)                 '配列の再定義
    For i = 1 To rsCnt
        DopeName(i) = rs("DOPKIND")
        
        If Len(DopeName(i)) < 7 Then         '7桁にあわせる
            DopeName(i) = DopeName(i) & Space(7 - Len(DopeName(i)))
        Else
            DopeName(i) = Left(DopeName(i), 7)
        End If
        
        rs.MoveNext
    Next
    rs.Close

    GetDopeNames = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :指定領域の品番を書き換える（品番管理Tbl対象）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型          ,説明
'          :CRYNUM        ,I  ,String      ,結晶番号
'          :ChgFrom       ,I  ,Integer     ,領域開始位置
'          :ChgLength     ,I  ,Integer     ,領域終了位置
'          :hin           ,I  ,tFullHinban ,書き換え後品番
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :
'履歴      :2001/08/11 作成  野村
Public Function ChangeAreaHinban(CRYNUM$, ChgFrom%, ChgLength%, HIN As tFullHinban) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim ChgTo As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function ChangeAreaHinban"

    ChangeAreaHinban = FUNCTION_RETURN_FAILURE
    ChgTo = ChgFrom + ChgLength
    
    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
Debug.Print "=== ChangeAreaHinban ==="
    ''指定領域を全て含む品番があれば指定領域の下側となるレコードを分割作成する
    sql = "insert into TBCME041 (CRYNUM,INGOTPOS,HINBAN,REVNUM,FACTORY,OPECOND,LENGTH,REGDATE,UPDDATE,SENDFLAG,SENDDATE) " & _
          "select CRYNUM, " & ChgTo & ", HINBAN, REVNUM, FACTORY, OPECOND, INGOTPOS+LENGTH-" & ChgTo & ", REGDATE, UPDDATE, SENDFLAG, SENDDATE " & _
          "From TBCME041 " & _
          "where (CRYNUM='" & CRYNUM & "') and (INGOTPOS<" & ChgFrom & ") and (INGOTPOS+LENGTH>" & ChgTo & ")"
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "全域を含む品番はなかった"
    Else
    ''     WriteDBLog sql
    End If
    
    ''指定領域の開始位置を含む品番があればそれを調整する
    sql = "update TBCME041 set LENGTH=" & ChgFrom & "-INGOTPOS, UPDDATE=SYSDATE " & _
          "where (CRYNUM='" & CRYNUM & "') and (INGOTPOS<" & ChgFrom & ") and (INGOTPOS+LENGTH>" & ChgFrom & ")"
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "開始位置を含む品番はなかった"
    Else
    ''    WriteDBLog sql
    End If
    
    ''指定領域の終了位置を含む品番があればそれを調整する
    sql = "update TBCME041 set INGOTPOS=" & ChgTo & ", LENGTH=INGOTPOS+LENGTH-" & ChgTo & ", UPDDATE=SYSDATE " & _
          "where (CRYNUM='" & CRYNUM & "') and (INGOTPOS<" & ChgTo & ") and (INGOTPOS+LENGTH>" & ChgTo & ")"
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "終了位置を含む品番はなかった"
    Else
    ''    WriteDBLog sql
    End If
    
    ''指定領域内に全域が含まれる品番を削除する(一致するレコードを含む)
    sql = "delete from TBCME041 where (CRYNUM='" & CRYNUM & "') and (INGOTPOS>=" & ChgFrom & ") and (INGOTPOS+LENGTH<=" & ChgTo & ")"
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "全域が含まれる品番はなかった"
    Else
    ''    WriteDBLog sql
    End If
    
    ''指定領域の品番を追加する
    With HIN
    ''    WriteDBLog sql
        sql = "insert into TBCME041 (CRYNUM,INGOTPOS,HINBAN,REVNUM,FACTORY,OPECOND,LENGTH,REGDATE,UPDDATE,SENDFLAG,SENDDATE) values " & _
              "('" & CRYNUM & "', " & ChgFrom & ", '" & .hinban & "', " & .mnorevno & ", '" & .factory & "', '" & .opecond & "'," & _
              ChgLength & ", SYSDATE, SYSDATE, '0', SYSDATE)"
    End With
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "品番追加失敗"
        GoTo proc_exit
    Else
    ''    WriteDBLog sql
    End If
    
    If dbIsMine Then
        OraDBClose
    End If
    
    ChangeAreaHinban = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :ブロックがホールド状態がどうか調べる。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :DopeName()    ,I  ,String    ,ブロックID
'          :戻り値        ,O  ,Integer   ,0:ホールド状態でない 1:ホールド状態 -1:読み込みエラー
'説明      :
'履歴      :2001/09/18 作成  蔵本
Public Function CheckHoldBlock(BLOCKID As String) As Integer

    Dim rs      As OraDynaset               '抽出RecordDynaset
    Dim sql     As String                   'SQL文

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function CheckHoldBlock"

    CheckHoldBlock = 0
    
    'SQL文の作成
    sql = "select HOLDCLS from TBCME040 where BLOCKID='" & BLOCKID & "' "
     
    'データの抽出
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '''抽出レコードが存在しない場合
    If rs.EOF Then
        CheckHoldBlock = -1
        GoTo proc_exit
    End If
            
    If rs("HOLDCLS") = 1 Then
        CheckHoldBlock = 1
    End If

    rs.Close

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    CheckHoldBlock = -1
    Resume proc_exit
End Function

'概要      :結晶の型を調べる
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Crynum        ,I  ,String    ,結晶番号
'          :戻り値        ,O  ,String    ,"P+","P-","N+","N-","Unknown" のいずれか
'説明      :
'履歴      :2002/03/28 作成  野村
Public Function GetXlType(CRYNUM$) As String
Dim sql As String
Dim rs As OraDynaset               '抽出RecordDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetXlType"
    
    GetXlType = "Unknown"

    sql = "select"
    sql = sql & " case when E018.HSXTYPE='P' then"
    sql = sql & "   case when E018.HSXRMAX <="
    sql = sql & "             (select to_number(INFO1) from TBCMB005"
    sql = sql & "              where SYSCLASS='LG' and CLASS='02' and CODE='P+')"
    sql = sql & "        then 'P+' else 'P-' end"
    sql = sql & "   when E018.HSXTYPE='N' then"
    sql = sql & " case when E018.HSXRMAX <="
    sql = sql & "             (select to_number(INFO1) from TBCMB005"
    sql = sql & "              where SYSCLASS='LG' and CLASS='02' and CODE='N+')"
    sql = sql & "        then 'N+' else 'N-' end"
    sql = sql & " else 'Unknown'"
    sql = sql & " end as TYPE "
    sql = sql & "from TBCME037 XL, TBCME018 E018 "
    sql = sql & "where E018.HINBAN=XL.RPHINBAN"
    sql = sql & "  and E018.MNOREVNO=XL.RPREVNUM"
    sql = sql & "  and E018.FACTORY=XL.RPFACT"
    sql = sql & "  and E018.OPECOND=XL.RPOPCOND"
    sql = sql & "  and XL.CRYNUM='" & CRYNUM & "'"

    'データの抽出
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        GetXlType = rs("TYPE")
    End If
    rs.Close
    Set rs = Nothing

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'==========================================
' 連番取得関数
'==========================================


'概要      :引上指示Noを取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :gouki         ,I  ,String    ,号機ID
'          :戻り値        ,O  ,String    ,新引上指示Noの連番部分
'説明      :
'履歴      :2001/06/20 作成  野村
Public Function GetNewID_Siji(GOUKI$) As String
Dim key As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010b.bas -- Function GetNewID_Siji"

    ''固定部を得る(号機+年度)
    key = GOUKI & (year(oraGetSysdate()) Mod 10)
    GetNewID_Siji = key & GetNewSeq(SEQ_HIKIAGE_SIJI, key) & "00"

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :リメルト原料番号の連番部分を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :戻り値        ,O  ,String    ,新リメルト原料番号の連番部分
'説明      :
'履歴      :2001/06/20 作成  野村
Public Function GetNewID_RemeltGenryo() As String
Dim key As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010b.bas -- Function GetNewID_RemeltGenryo"

    ''固定部はないため、「_」とする
    key = "_"
    GetNewID_RemeltGenryo = GetNewSeq(SEQ_RMLT_GENRYO, key)

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :結晶番号の連番部分を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :gouki         ,I  ,String    ,号機ID
'          :戻り値        ,O  ,String    ,新結晶番号の連番部分
'説明      :
'履歴      :2001/06/20 作成  野村
Public Function GetNewID_CryNum(GOUKI$) As String
Dim key As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010b.bas -- Function GetNewID_CryNum"

    ''固定部を得る(号機+年度)
    key = GOUKI & (year(oraGetSysdate()) Mod 10)
    GetNewID_CryNum = GetNewSeq(SEQ_CRYNUM, key)

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :サンプルNoを取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :戻り値        ,O  ,String    ,新サンプルNo
'説明      :
'履歴      :2001/06/20 作成  野村
Public Function GetNewID_SampleNo() As String
Dim key As String
Dim sql As String
Dim rs As OraDynaset
Dim newID As String
Dim firstID As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010b.bas -- Function GetNewID_SampleNo"
    
    GetNewID_SampleNo = 0

    ''固定部はないため、「_」とする
    key = "_"
    newID = GetNewSeq(SEQ_SAMPLENO, key)
    
'>>>>> サンプル№6桁対応 2007/05/25 SETsw kubota -------------
    newID = SAMPLENO_HEAD & Format$(newID, "00000")
'<<<<< サンプル№6桁対応 2007/05/25 SETsw kubota -------------
    
    firstID = newID
    Do
        sql = "select REPSMPLIDCS from XSDCS where (REPSMPLIDCS='" & newID & "') and (KTKBNCS='0')"
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            '該当なしなら、その番号を使ってよい
            rs.Close
            Exit Do
        End If
        rs.Close
        
        newID = GetNewSeq(SEQ_SAMPLENO, key)

'>>>>> サンプル№6桁対応 2007/05/25 SETsw kubota -------------
        newID = SAMPLENO_HEAD & Format$(newID, "00000")
'<<<<< サンプル№6桁対応 2007/05/25 SETsw kubota -------------
        
        If newID = firstID Then
            'まずないはずだが、１周全て生きたサンプルだった場合
            newID = 0
            Exit Do
        End If
    Loop

    GetNewID_SampleNo = newID

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :新たな連番を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SeqCode       ,I  ,String    ,連番種別管理コード
'          :key           ,I  ,String    ,連番種別コード
'          :戻り値        ,O  ,String    ,連番文字列
'説明      :
'履歴      :2001/06/20 作成  野村
Private Function GetNewSeq(SeqCode$, key$) As String
Dim rs As OraDynaset
Dim sql As String
Dim seq As Long
Dim keta As Integer
Dim clrWhen As String
Dim clrAt As Date
Dim sysNow As Date
Dim dbIsMine As Boolean

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010b.bas -- Function GetNewSeq"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    
    ''連番管理 から、指定品番の仕様レコードを新しい順に取り出す
    sql = "select CONTNUM, MAXFIG, NUMUNIT, CLRDATE from TBCMB015 where (CNTMNGCD='" & SeqCode & "') and (CNTNUMCD='" & key & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、1番を登録して返す
        seq = 1                     '新番号 = 1
        rs.Close
        
        ''桁数を得る
        sql = "select MAXFIG from tbcmb015 where (cntmngcd='" & SeqCode & "') and (cntnumcd='DEF')"
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Debug.Print "GetNewSeq: 連番管理TBLに登録されていない (" & SeqCode & ")"
            keta = 3    '既定値は3桁としておく。
        Else
            keta = rs("MAXFIG")         '桁数
        End If
        rs.Close
        
        ''連番管理テーブルに指定キーの行を追加する
        sql = "insert into tbcmb015 (cntmngcd,cntnumcd,contnum,maxfig,numunit,numname,clrdate,regdate,upddate) " & _
              "(select cntmngcd, '" & key & "', 1, " & keta & ", numunit, numname, sysdate, sysdate, sysdate" & _
              " From TBCMB015 where (cntmngcd='" & SeqCode & "') and (cntnumcd='DEF'))"
        OraDB.ExecuteSQL sql
    Else
        ''見つかったら、連番を1つ上げる
        seq = rs("CONTNUM") + 1     '現在の番号+1
        keta = rs("MAXFIG")         '桁数
        clrWhen = rs("NUMUNIT")     'クリアタイミング
        clrAt = rs("CLRDATE")       '前回クリアした日時
        rs.Close
        
        ''桁数オーバーになったら連番を1に戻す
        If Len(CStr(seq)) > keta Then
            seq = 1
        End If
        
        ''クリア時期が来ていたら、連番を1に戻す
        sysNow = oraGetSysdate()
        Select Case clrWhen
          Case "Y"      '年次クリア
                If year(clrAt) <> year(sysNow) Then
                    seq = 1
                End If
          Case "M"      '月次クリア
                If (year(clrAt) <> year(sysNow)) Or (month(clrAt) <> month(sysNow)) Then
                    seq = 1
                End If
          Case "D"
                If (year(clrAt) <> year(sysNow)) Or (month(clrAt) <> month(sysNow)) Or (day(clrAt) <> day(sysNow)) Then
                    seq = 1
                End If
        End Select
        
        ''連番管理テーブルを更新する
        sql = "update tbcmb015 set" & _
              " contnum=" & seq & _
              ",upddate=sysdate"
        If seq = 1 Then     'クリアした場合
            sql = sql & ",clrdate=sysdate"
        End If
        sql = sql & " where (cntmngcd='" & SeqCode & "') and (cntnumcd='" & key & "')"
        OraDB.ExecuteSQL sql
    End If
    
    If dbIsMine Then
        OraDBClose
    End If
    
    GetNewSeq = Format$(seq, String(keta, "0"))

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "==== ERROR SQL ===="
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :フォームを画面中央に移動
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :frmObj        ,I   ,Form      ,フォームオブジェクト
'説明      :
Public Sub CenterForm(frmObj As Form)
    '' フォームを画面中央に移動

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub CenterForm"

    With frmObj
        If .WindowState <> 2 Then
            .Left = (Screen.Width - .Width) / 2
            .Top = (Screen.Height - .Height) / 2
        End If
    End With
    

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :コントロールオブジェクトに現在時刻をセットする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :ctrlObj       ,I   ,Control   ,コントロールオブジェクト
'説明      :
Public Sub SetPresentTime(ctrlObj As Control)
    '' 現在時刻の取得とセット

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SetPresentTime"

    ctrlObj.Caption = Format$(Now, "yyyy/mm/dd hh:nn")

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


'概要      :フォーム表示処理（次画面移動処理）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :frmOwnerObj   ,I   ,Form      ,オーナーフォームオブジェクト（呼び出し元）
'          :frmShowObj    ,I   ,Form      ,表示フォームオブジェクト
'説明      :
Public Sub ShowFormProc(frmOwnerObj As Form, frmShowObj As Form)

    '' マウスカーソルを処理中に変更

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub ShowFormProc"

    Screen.MousePointer = vbHourglass
    '' フォームの表示
    frmShowObj.Show
    '' オーナー画面を隠す
    frmOwnerObj.Hide
    '' マウスカーソルを矢印に戻す
    Screen.MousePointer = vbDefault


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub



'概要      :フォームクローズ処理(前画面戻処理)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :frmPrevObj    ,I   ,Form      ,前表示フォームオブジェクト（閉じた後に表示させるフォーム）
'          :frmCloseObj   ,I   ,Form      ,クローズフォームオブジェクト（閉じるフォーム）
'説明      :
Public Sub CloseFormProc(frmPrevObj As Form, frmCloseObj As Form)

    '' マウスカーソルを処理中に変更

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub CloseFormProc"

    Screen.MousePointer = vbHourglass
    '' フォームをクローズする
    Unload frmCloseObj
    DoEvents
    '' 前画面を表示する
    frmPrevObj.Show
    DoEvents
    '' マウスカーソルを矢印に戻す
    Screen.MousePointer = vbDefault
    DoEvents


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


'概要      :処理開始処理
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :[strMsg]      ,I   ,String    ,表示メッセージ文字列
'説明      :時間のかかる処理を行うときの前処理を行う
'           EndProcess()と併用する。
Public Sub BeginProcess(Optional strMsg As String = "")

    '' メッセージがある場合、メッセージを表示する

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub BeginProcess"

    If strMsg <> "" Then
        MsgBox strMsg, vbOKOnly + vbInformation
    End If

    '' マウスカーソルを処理中に変更
    Screen.MousePointer = vbHourglass
    DoEvents

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :処理終了処理
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :[strMsg]      ,I   ,String    ,表示メッセージ文字列
'説明      :時間のかかる処理を行った後の後処理を行う。
'           BeginProcess()と併用する。
Public Sub EndProcess(Optional strMsg As String = "")

    '' メッセージがある場合、メッセージを表示する

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub EndProcess"

    If strMsg <> "" Then
        MsgBox strMsg, vbOKOnly + vbInformation
    
    End If

    '' マウスカーソルを矢印に戻す
    Screen.MousePointer = vbDefault
    DoEvents

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub





'概要      :スプレッドコントロールの初期化処理
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :ctrlObj       ,I   ,vaSpread   ,スプレッドコントロールオブジェクト
'          :[lMaxRows]    ,I   ,Long      ,スプレッドの初期表示行数
'説明      :
Public Sub SpCtrlInit(ctrlObj As vaSpread, Optional lMaxRows As Long = -1)
    
    '' スプレッドの初期表示行数をセット

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlInit"

    If lMaxRows >= 0 Then
        ''　初期表示行数指定がある場合、スプレッドの初期表示行数をセット
        ctrlObj.MaxRows = lMaxRows
    End If
    '' セルのロックを反映する
    ctrlObj.Protect = True


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


'概要      :スプレッドに行追加
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :ctrlObj       ,I   ,vaSpread   ,スプレッドコントロールオブジェクト
'説明      :
Public Sub SpCtrlInsertRow(ctrlObj As vaSpread)

    Dim lSmpPos As Long


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlInsertRow"

    ctrlObj.MaxRows = ctrlObj.MaxRows + 1
    lSmpPos = ctrlObj.MaxRows
    
    With ctrlObj
        .row = lSmpPos
        .row2 = lSmpPos + 1
        .BlockMode = True
        .Action = ActionInsertRow
        .BlockMode = False
    End With

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub




'概要      :コントロールの状態を設定する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :ctrlObj       ,I   ,Control       ,コントロールオブジェクト
'          :ctrlState     ,I   ,enm_CtrlStateKind ,コントロールの状態指示
'          :[bClear]      ,I   ,Boolean       ,コントロールテキスト内容のクリア指示（True：クリア False：クリアしない）
'説明      :
Public Sub CtrlEnabled(ctrlObj As Control, ctrlState As enm_CtrlStateKind, Optional bClear As Boolean = False)


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub CtrlEnabled"

    On Error Resume Next
    
    If TypeOf ctrlObj Is Frame Then      '' コントロールがフレームの場合
        '' コントロール指定状態をチェック
        Select Case ctrlState
            Case CTRL_DISABLE             '' 編集不可の場合
                ctrlObj.Enabled = False         '' オブジェクトを使用不可にする
            Case Else                       '' その他の場合
                ctrlObj.Enabled = True          '' オブジェクトを使用可能にする
        End Select
    Else                                    '' コントロールがフレーム以外の場合
        '' コントロール指定状態をチェック
        Select Case ctrlState
            Case CTRL_DISABLE             '' 編集不可の場合
                ctrlObj.BackColor = COLOR_DISABLE    '' 背景色を表示項目色にする
                ctrlObj.Locked = True                   '' ロックする
                ctrlObj.TabStop = False                 '' タブストップしない
            Case CTRL_DISABLE_GRAY        '' 編集不可(グレー色表示)の場合
                ctrlObj.BackColor = COLOR_GRAY '' 背景色をグレー色にする
                ctrlObj.Locked = True                   '' ロックする
                ctrlObj.TabStop = False                 '' タブストップしない
            Case CTRL_WARNING               '' 警告指示の場合
                ctrlObj.BackColor = COLOR_WARNING      '' 背景色を警告色にする
                ctrlObj.Locked = False                  '' ロックしない
                ctrlObj.TabStop = True                  '' タブストップする
            Case CTRL_DISABLE_WARNING               '' 警告指示編集不可の場合
                ctrlObj.BackColor = COLOR_WARNING      '' 背景色を警告色にする
                ctrlObj.Locked = True                  '' ロックする
                ctrlObj.TabStop = False                 '' タブストップしない
            Case CTRL_SELECTED
                ctrlObj.BackColor = COLOR_SELECTED    '' 背景色を選択色にする
                ctrlObj.Locked = True                   '' ロックする
                ctrlObj.TabStop = False                 '' タブストップしない
            Case CTRL_ENABLE_YELLOW
                ctrlObj.BackColor = COLOR_YELLOW        '' 背景色をイエローにする
                ctrlObj.Locked = False                   '' ロックしない
                ctrlObj.TabStop = True                   '' タブストップする
            Case CTRL_DISABLE_SKY
                ctrlObj.BackColor = COLOR_SKY      '' 背景色を警告色にする
                ctrlObj.Locked = True                  '' ロックする
                ctrlObj.TabStop = False                 '' タブストップしない
            Case Else                       '' その他の状態指定の場合
                ctrlObj.BackColor = COLOR_ENABLE        '' 背景色をウインドウの背景色にする
                ctrlObj.Locked = False                   '' ロックしない
                ctrlObj.TabStop = True                   '' タブストップする
        End Select
    
        ''テキストクリアチェック
        If bClear = True Then
            ctrlObj = ""
        End If
        
    End If
    
    '' コントロール状態指示が警告指示の場合
    Select Case ctrlState
    Case CTRL_WARNING
        ctrlObj.SetFocus     '' フォーカスをセットする
    End Select
    
    On Error GoTo 0


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


'概要      :スプレッドコントロールのセルの状態を設定する（単一セル）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                ,説明
'          :ctrlObj       ,   ,Control           ,
'          :Col           ,   ,Long              ,
'          :Row           ,   ,Long              ,
'          :ctrlState     ,   ,enm_CtrlStateKind ,
'          :[bClear]      ,   ,Boolean           ,
'説明      :
Public Sub SpCtrlEnabled(ctrlObj As Control, ByVal col As Long, ByVal row As Long, ctrlState As enm_CtrlStateKind, Optional bClear As Boolean = False)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlEnabled"

    SpCtrlBlockEnabled ctrlObj, col, row, col, row, ctrlState, bClear

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


'概要      :スプレッドコントロールのセルの状態を設定する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                ,説明
'          :ctrlObj       ,I   ,vaSpread           ,スプレッドコントロールオブジェクト
'          :Col           ,I   ,Long              ,列　Min位置（範囲指定）
'          :Row           ,I   ,Long              ,行　Min位置（範囲指定）
'          :Col2          ,I   ,Long              ,列　Max位置（範囲指定）
'          :Row2          ,I   ,Long              ,行　Max位置（範囲指定）
'          :ctrlState     ,I   ,enm_CtrlStateKind ,コントロールの状態指示
'          :[bClear]      ,I   ,Boolean           ,コントロールテキスト内容のクリア指示（True：クリア False：クリアしない）
'説明      :
Public Sub SpCtrlBlockEnabled(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, ByVal col2 As Long, ByVal row2 As Long, ctrlState As enm_CtrlStateKind, Optional bClear As Boolean = False)


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlBlockEnabled"

    On Error Resume Next

    ctrlObj.col = col
    ctrlObj.row = row
    ctrlObj.col2 = col2
    ctrlObj.row2 = row2
    
    '' スプレッドコントロールのセルを指定した状態にセットする。
    ctrlObj.BlockMode = True
    Select Case ctrlState
        Case CTRL_DISABLE         '' 編集不可の場合
            ctrlObj.BackColor = COLOR_DISABLE           '' 背景色を表示項目色にする
            ctrlObj.Lock = True                         '' ロックする
        Case CTRL_DISABLE_GRAY    '' 編集不可(グレー色表示)の場合
            ctrlObj.BackColor = COLOR_GRAY      '' 背景色をグレー色にする
            ctrlObj.Lock = True                         '' ロックする
        Case CTRL_DISABLE_SKY    '' 編集不可(推定値)の場合
            ctrlObj.BackColor = COLOR_SKY      '' 背景色を推定色にする
            ctrlObj.Lock = True                         '' ロックする
        'Add Start 2010/08/04 SMPK Nakamura
        Case CTRL_ENABLE_SKY     '' 編集可(推定値)の場合
            ctrlObj.BackColor = COLOR_SKY      '' 背景色を推定色にする
            ctrlObj.Lock = False                         '' ロックする
        'Add End 2010/08/04 SMPK Nakamura
        Case CTRL_WARNING           '' 警告指示の場合
            ctrlObj.BackColor = COLOR_WARNING           '' 背景色を赤色表示にする
            ctrlObj.Lock = False                        '' ロックしない
        Case CTRL_DISABLE_WARNING           '' 警告指示編集不可の場合
            ctrlObj.BackColor = COLOR_WARNING           '' 背景色を赤色表示にする
            ctrlObj.Lock = True                        '' ロックする
        Case CTRL_SELECTED
            ctrlObj.BackColor = COLOR_SELECTED          '' 背景色を選択色にする
            ctrlObj.Lock = True                         '' ロックする
        Case CTRL_ENABLE_GRAY    '' 編集可(グレー色表示)の場合
            ctrlObj.BackColor = COLOR_GRAY      '' 背景色をグレー色にする
            ctrlObj.Lock = False                        '' ロックしない
        Case CTRL_ENABLE_YELLOW
            ctrlObj.BackColor = COLOR_YELLOW        '' 背景色をイエローにする
            ctrlObj.Lock = False                   '' ロックしない
        Case CTRL_DISABLE_YELLOW
            ctrlObj.BackColor = COLOR_YELLOW        '' 背景色をイエローにする
            ctrlObj.Lock = True                   '' ロックする
        '------ kuramoto 追加 2001/09/25 ------
        Case CTRL_ENABLE_RED
            ctrlObj.BackColor = COLOR_RED        '' 背景色をレッドにする
            ctrlObj.Lock = False                  '' ロックしない
        Case CTRL_DISABLE_RED
            ctrlObj.BackColor = COLOR_RED        '' 背景色をレッドにする
            ctrlObj.Lock = True                  '' ロックする
        '--------------------------------------
        Case Else                   '' その他の状態指定の場合
            ctrlObj.BackColor = COLOR_ENABLE            '' 背景色を白色表示にする
            ctrlObj.Lock = False                        '' ロックしない
    End Select
    ctrlObj.BlockMode = False

    ''テキストクリアチェック
    If bClear = True Then
        Dim iCol As Long
        Dim IRow As Long
        For IRow = row To row2
            For iCol = col To col2
                ctrlObj.SetText iCol, IRow, ""
            Next iCol
        Next IRow
    End If
    
    '' セルのロックを反映する
    ctrlObj.Protect = True
    
    
    '' 処理対象サンプルをアクティブセルにする
    Select Case ctrlState
        Case CTRL_WARNING           '' 警告指示の場合
        SpCtrlSetAction ctrlObj, col, row, col2, row2, ActionSelectBlock
    End Select
    
    On Error GoTo 0
    

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


'概要      :スプレッドにコンボボックスを設定する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :ctrlObj       ,I   ,vaSpread  ,スプレッドコントロールオブジェクト
'          :Col           ,I   ,Long     ,列　Min位置（範囲指定）
'          :Row           ,I   ,Long     ,行　Min位置（範囲指定）
'          :Col2          ,I   ,Long     ,列　Max位置（範囲指定）
'          :Row2          ,I   ,Long     ,行　Max位置（範囲指定）
'          :strItem       ,I   ,String   ,コンボボックス表示項目内容文字列（Tab区切り）
'          :[lParam]      ,I   ,Long     ,コンボボックス初期表示項目指定
'説明      :
Public Sub SpCtrlSetCombo(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, strItem As String, Optional lParam As Long = 0)

    Dim x As Long
    Dim Y As Long


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlSetCombo"

    On Error Resume Next

    ctrlObj.col = col
    ctrlObj.row = row
    
    ctrlObj.CellType = CellTypeComboBox
    ctrlObj.TypeComboBoxList = strItem
    ctrlObj.TypeComboBoxCurSel = lParam
    
    On Error GoTo 0


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :指定した動作のスプレッド処理を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :ctrlObj       ,I   ,vaSpread  ,スプレッドコントロールオブジェクト
'          :Col           ,I   ,Long     ,列　Min位置（範囲指定）
'          :Row           ,I   ,Long     ,行　Min位置（範囲指定）
'          :Col2          ,I   ,Long     ,列　Max位置（範囲指定）
'          :Row2          ,I   ,Long     ,行　Max位置（範囲指定）
'          :iAction       ,I   ,Integer  ,スプレッド処理動作指示
'説明      :
Public Sub SpCtrlSetAction(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, ByVal col2 As Long, ByVal row2 As Long, iAction As Integer)


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlSetAction"

    On Error Resume Next

    '' 指定した動作のスプレッド処理を行う
    ctrlObj.col = col
    ctrlObj.row = row
    ctrlObj.col2 = col2
    ctrlObj.row2 = row2
    ctrlObj.BlockMode = True
    ctrlObj.Action = iAction
    ctrlObj.BlockMode = False
    
    On Error GoTo 0


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :指定セルのロック状態を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :ctrlObj       ,I   ,Control  ,スプレッドコントロールオブジェクト
'          :Col           ,I   ,Long     ,列　Min位置（範囲指定）
'          :Row           ,I   ,Long     ,行　Min位置（範囲指定）
'          :戻り値        ,O  ,Boolean   ,True:ロック     False:ロックされていない
'説明      :
Public Function SpCtrlIsLock(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long) As Boolean


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Function SpCtrlIsLock"

    ctrlObj.col = col
    ctrlObj.row = row

    SpCtrlIsLock = ctrlObj.Lock


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :指定セルをマーキングする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                ,説明
'          :ctrlObj       ,I   ,vaSpread           ,スプレッドコントロールオブジェクト
'          :Col           ,I   ,Long              ,列　Min位置（範囲指定）
'          :Row           ,I   ,Long              ,行　Min位置（範囲指定）
'          :Col2          ,I   ,Long              ,列　Max位置（範囲指定）
'          :Row2          ,I   ,Long              ,行　Max位置（範囲指定）
'説明      :
Public Sub SpCtrlSetMark(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, ByVal col2 As Long, ByVal row2 As Long)
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlSetMark"

    On Error Resume Next

    ctrlObj.col = col
    ctrlObj.row = row
    ctrlObj.col2 = col2
    ctrlObj.row2 = row2
    
    '' スプレッドコントロールのセルを指定した状態にセットする。
    ctrlObj.BlockMode = True
    ctrlObj.BackColor = &HFFFF80       '' 空色
    ctrlObj.BlockMode = False
   
    '' セルのロックを反映する
    ctrlObj.Protect = True
    
    On Error GoTo 0

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


'概要      :指定セルのロックを設定・解除する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                ,説明
'          :ctrlObj       ,I   ,vaSpread           ,スプレッドコントロールオブジェクト
'          :Col           ,I   ,Long              ,列　Min位置（範囲指定）
'          :Row           ,I   ,Long              ,行　Min位置（範囲指定）
'          :Col2          ,I   ,Long              ,列　Max位置（範囲指定）
'          :Row2          ,I   ,Long              ,行　Max位置（範囲指定）
'          :[bLock]       ,I   ,Boolean           ,コントロールの状態指示(True：ロック False：ロックしない)
'          :[bClear]      ,I   ,Boolean           ,コントロールテキスト内容のクリア指示（True：クリア False：クリアしない）
'説明      :
Public Sub SpCtrlSetLock(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, ByVal col2 As Long, ByVal row2 As Long, Optional Block As Boolean = True, Optional bClear As Boolean = False)
    
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlSetLock"

    On Error Resume Next

    ctrlObj.col = col
    ctrlObj.row = row
    ctrlObj.col2 = col2
    ctrlObj.row2 = row2
    
    '' スプレッドコントロールのセルを指定した状態にセットする。
    ctrlObj.BlockMode = True
    ctrlObj.Lock = Block
    ctrlObj.BlockMode = False

    ''テキストクリアチェック
    If bClear = True Then
        Dim iCol As Long
        Dim IRow As Long
        For IRow = row To row2
            For iCol = col To col2
                ctrlObj.SetText iCol, IRow, ""
            Next iCol
        Next IRow
    End If
    
    '' セルのロックを反映する
    ctrlObj.Protect = True
    
    On Error GoTo 0

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :指定セルのフォントを太字にする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                ,説明
'          :ctrlObj       ,I   ,vaSpread           ,スプレッドコントロールオブジェクト
'          :Col           ,I   ,Long              ,列　Min位置（範囲指定）
'          :Row           ,I   ,Long              ,行　Min位置（範囲指定）
'          :Col2          ,I   ,Long              ,列　Max位置（範囲指定）
'          :Row2          ,I   ,Long              ,行　Max位置（範囲指定）
'          :[bState]      ,I   ,Boolean           ,文字状態指示（True:太字指定あり False:太字指定なし）
'説明      :
Public Sub SpCtrlFontBold(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, ByVal col2 As Long, ByVal row2 As Long, Optional bState As Boolean = False)
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlFontBold"

    On Error Resume Next
    
    ctrlObj.col = col
    ctrlObj.row = row
    ctrlObj.col2 = col2
    ctrlObj.row2 = row2
    ctrlObj.BlockMode = True
    ctrlObj.FontBold = bState
    ctrlObj.BlockMode = False
    On Error GoTo 0

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub



'概要      :フォームの全テキストボックスについて、.TextをRTrimする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :frm           ,I  ,Form      ,対象フォーム
'説明      :
'履歴      :2001/08/24 作成  野村
Public Sub TrimAll(frm As Form)
Dim ctl As Control

    For Each ctl In frm.Controls
        If TypeName(ctl) = "TextBox" Then
            ctl.Text = RTrim$(ctl.Text)
        End If
    Next
End Sub


'概要      :コントロールコードをスペースに置き換える
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :s             ,I  ,String    ,元文字列
'説明      :DebugおよびSQL異常の対処用
'履歴      :2001/09/26 作成  野村
Public Function toNormalStr(s$) As String
Dim i As Integer

    On Error Resume Next
    For i = 1 To Len(s)
        If Asc(Mid$(s, i, 1)) < &H20 Then
            Debug.Print "toNormalStr(""" & s & """) : " & i & "文字目=&H" & Asc(Mid$(s, i, 1))
            Mid$(s, i, 1) = " "
        End If
    Next
    toNormalStr = s
End Function


'概要      :抵抗値表示用の小数部桁数を求める
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :rs            ,I  ,Double    ,抵抗値
'          :[IgnoreZero]  ,I  ,Boolean   ,0が与えられたときに1を返す（False:5を返す)
'説明      :抵抗値表示桁数を統一するため。1～5桁の、小数表示桁数を求める
'履歴      :2002/1/15 作成  野村
Public Function GetLowerCol(ByVal rs As Double, Optional IgnoreZero As Boolean = False) As Integer
    rs = Abs(rs)
    If rs = 0 Then
        If IgnoreZero Then
            '0を無効入力とみなして１を返す
            GetLowerCol = 1
        Else
            '0を数値とみなして５を返す
            GetLowerCol = 5
        End If
    ElseIf rs >= 10000 Then
        GetLowerCol = 1
    ElseIf rs >= 1000 Then
        GetLowerCol = 2
    ElseIf rs >= 100 Then
        GetLowerCol = 3
    ElseIf rs >= 10 Then
        GetLowerCol = 4
    ElseIf rs > 0 Then
        GetLowerCol = 5
    Else
        'マイナス値は入ってこないはず
        GetLowerCol = 0
    End If
End Function



'概要      :抵抗の値を表示用に文字列化する(有効6桁+小数点)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :rs_1          ,I  ,Double    ,抵抗値1
'          :[rs_2]        ,I  ,Double    ,抵抗値2
'説明      :抵抗値表示桁数を統一するため。抵抗値2を入れると、範囲文字列を返す
'履歴      :2001/12/21 作成  野村
Public Function toRsStr(rs_1 As Double, Optional rs_2 As Double = -1#) As String
Dim s$, rsStr$

    If rs_1 >= 99999.9 Then
        s = "99999.9"
    Else
        s = Format$(rs_1, "0." & String(GetLowerCol(rs_1), "0"))
    End If
    rsStr = s
    
    If rs_2 >= 0 Then
        rsStr = rsStr & "-"
        
        If rs_2 >= 99999.9 Then
            s = "99999.9"
        Else
            s = Format$(rs_2, "0." & String(GetLowerCol(rs_2), "0"))
        End If
        rsStr = rsStr & s
    End If
    
    toRsStr = rsStr
End Function

'概要      :抵抗の値を表示用に文字列化する(指定の小数点以下桁数)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :rs            ,I  ,Double    ,抵抗値
'          :place         ,I  ,Integer   ,小数点以下桁数
'説明      :抵抗値表示桁数を統一するため。<0のときは空文字列を返す
'履歴      :2002/1/16 作成  野村
'履歴      :2002/1/17 S.Sano
Public Function toRsStrByPlace(rs As Double, place As Integer) As String
Dim s$

    If rs < 0 Then
        s = vbNullString
'2002/01/17 S.Sano    ElseIf rs >= 99999.9 Then
'2002/01/17 S.Sano        s = "99999.9"
    Else
        s = Format$(rs, "0." & String(place, "0"))
        If val(s) >= 100000 Then
            s = "99999." & String(place, "9")
        End If
    End If
    toRsStrByPlace = s
End Function

Public Function toRsStr_nl(rs_1 As Double, Optional rs_2 As Double = -1#) As String '抵抗の表示 2003/12/8
Dim s$, rsStr$
    
    If rs_1 < 0 Then  '-1(Null)のとき
        s = vbNullString
    Else
        If rs_1 >= 99999.9 Then
            s = "99999.9"
        Else
            s = Format$(rs_1, "0." & String(GetLowerCol(rs_1), "0"))
        End If
    End If
            rsStr = s
    
    If rs_2 >= 0 Then
        rsStr = rsStr & "-"
        
        If rs_2 >= 99999.9 Then
            s = "99999.9"
        Else
            s = Format$(rs_2, "0." & String(GetLowerCol(rs_2), "0"))
        End If
        rsStr = rsStr & s
    Else    'rs_2が-1(Null)のときの処理
        rsStr = rsStr & "-"
        s = vbNullString
        rsStr = rsStr & s
    End If
    
    toRsStr_nl = rsStr
End Function

'概要      :スプレッドの抵抗実数表示を仕様に合わせて桁を揃える。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :targetSpread  ,I  ,vaSpread  ,対象スプレッド
'          :col1    ,I  ,Long      ,対象カラム(From)
'          :[col2]  ,I  ,Long      ,対象カラム(To)
'          :[row1]  ,I  ,Long      ,対象行(From)
'          :[row2]  ,I  ,Long      ,対象行(To)
'説明      :スプレッドの対象範囲について、全てのセルで小数点以下桁数を揃える
'履歴      :2002/01/15 作成 S.Sano
'          :2002/01/16 修正 野村
Public Sub RsSpreadSet(targetSpread As vaSpread, col1 As Long, Optional col2 As Long = 0, Optional row1 As Long = 0, Optional row2 As Long = 0)
Dim row As Long
Dim col As Long
Dim MaxLowerCol As Integer
Dim lowCol As Integer
Dim rs As Double
    
    MaxLowerCol = 0
    '既定範囲の設定
    If row1 = 0 Then
        row1 = 1
        If row2 = 0 Then row2 = targetSpread.MaxRows
    ElseIf row2 = 0 Then
        row2 = row1
    End If
    If col2 = 0 Then
        col2 = col1
    End If
    
    With targetSpread
        .ReDraw = False
        
        '表示すべき小数点以下桁数を求める
        For col = col1 To col2
            For row = row1 To row2
                .GetFloat col, row, rs
                lowCol = GetLowerCol(rs, True)
                If MaxLowerCol < lowCol Then
                    MaxLowerCol = lowCol
                End If
            Next
        Next
        
        '小数点以下桁数を揃える
        .BlockMode = True
        .col = col1
        .col2 = col2
        .row = row1
        .row2 = row2
        .CellType = CellTypeFloat
        .TypeFloatMax = 99999.99999
        .TypeFloatMin = 0#
        '.TypeFloatMax = Val("99999." & Left("99999", MaxLowerCol))
        '.TypeFloatMin = 0
        .TypeFloatDecimalPlaces = MaxLowerCol
        .BlockMode = False
        
        .ReDraw = True
    End With
End Sub

Public Sub WriteDBLog(ByVal sqlStr$, Optional ByVal memo$ = " ")
Dim dbIsMine    As Boolean
Dim sql         As String
Dim s           As String
Dim i           As Integer
Dim hostname    As String
Dim fncName     As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    fncName = gErr.fncName
    gErr.Push "s_cmzc004c.bas -- Function WriteDBLog"

    ' 本来無いはずの関数名取得不可に対応
    ' ※エラーに関するPush、Popの処理回数不一致で発生する可能性有り
    If Trim(fncName) = "" Then
        fncName = "fncName is Nothing"
    ElseIf fncName = vbNullString Then
        fncName = "fncName is Null"
    End If

#If DBG Then
Dim fno As Integer
    fno = FreeFile
    Open App.Path & "\" & App.EXENAME & ".LOG" For Append As fno
    Print #fno, Now, fncName, memo
    Print #fno, "    " & sqlStr
    Close fno
#End If

    ''与えられたSQL中のシングルクォートを置き換える
    s = Replace(sqlStr, "'", "''")
    
    ''ホスト名を得る
    hostname = String(51, " ")
    GetComputerName hostname, 50
    If InStr(1, hostname, vbNullChar) Then
        hostname = Trim$(Left$(hostname, InStr(1, hostname, vbNullChar) - 1))
    End If

    ''SQLを作成
    If sqlStr = vbNullString Then sqlStr = " "
    If memo = vbNullString Then memo = " "
    sql = "insert into TBCMC003 " & _
        "(L_DATE, SEQ, HOSTNAME, APPNAME, FNCNAME, SQL, MEMO) values ("
    sql = sql & "sysdate, "                 'タイムスタンプ
    sql = sql & "LOG_SEQ.NEXTVAL, "         'SEQ
    sql = sql & "'" & hostname & "', "      '端末名
    sql = sql & "'" & App.EXENAME & "', "   'APPNAME
    sql = sql & "'" & fncName & "', "       '関数名
    sql = sql & "'" & s & "', "             'SQL
    sql = sql & "'" & memo & "' "           'memo
    sql = sql & ")"
    
    ''Logを書く
    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    OraDB.ExecuteSQL sql
    If dbIsMine Then
        OraDBClose
    End If
    
proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :指定領域の品番を書き換える（品番管理Tbl対象）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型          ,説明
'          :BLOCKID       ,I  ,String      ,結晶番号
'          :hin           ,I  ,tFullHinban ,書き換え後品番
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :
'履歴      :2003/10/29 二渡
Public Function ChangeXSDCSHinban(BLOCKID$, HIN As tFullHinban) As FUNCTION_RETURN
Dim rs As OraDynaset
Dim sql As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function ChangeAreaHinban"

    ChangeXSDCSHinban = FUNCTION_RETURN_FAILURE
          
    ''指定領域の開始位置を含む品番があればそれを調整する
    sql = "update XSDCS "
    sql = sql & "set HINBCS = '" & HIN.hinban & "',"
    sql = sql & "REVNUMCS = '" & HIN.mnorevno & "',"
    sql = sql & "FACTORYCS = '" & HIN.factory & "',"
'    sql = sql & "OPECS = " & HIN.OPECOND & ","
    sql = sql & "OPECS = '" & HIN.opecond & "',"    'ｼﾝｸﾞﾙｸｫｰﾄ追加 2009/11/16 SETsw Nakada
    sql = sql & "KDAYCS = sysdate,"
    sql = sql & "KSTAFFCS = '" & STAFFIDBUFF & "'"
    sql = sql & " WHERE LIVKCS = '0' AND "
    sql = sql & "CRYNUMCS = '" & BLOCKID$ & "'"
    
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "該当するブロックが無かった"
    Else
    ''    WriteDBLog sql
    End If
    
      
    ChangeXSDCSHinban = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :コンボボックスに結晶操業マスタ内の選択肢を設定する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SYSCLASS      ,I  ,String    ,SYS区分
'          :Class         ,I  ,String    ,区分
'          :FieldName     ,I  ,String    ,選択肢名称フィールド名
'          :cmb           ,O  ,ComboBox  ,設定先のコンボボックス
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :
'履歴      :2005/06/02 KODA9
Public Function SetCodeComboA9(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$, cmb As ComboBox) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer

    ''コードマスタから、指定のフィールドを引く
    sql = "SELECT CODEA9 , " & FieldName & " from KODA9" & _
          " WHERE (SYSCA9='" & SYSCLASS & "') and (SHUCA9='" & Class & "')" & _
          " Order by CTR01A9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、NG
        cmb.Clear
        SetCodeComboA9 = FUNCTION_RETURN_FAILURE
    Else
        ''見つかったら、コンボボックスに選択肢を設定する
        With cmb
            .Clear
            max = rs.RecordCount
            For i = 1 To max
                .AddItem GetGPCodeDspStr(rs("CODEA9"), rs(FieldName))
                rs.MoveNext
            Next
        End With
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

End Function
'概要      :結晶操業用コードを検索し、指定フィールドの内容を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SYSCLASS      ,I  ,String    ,システム区分
'          :CLASS         ,I  ,String    ,区分
'          :CODE          ,I  ,String    ,コード
'          :FieldName     ,I  ,String    ,フィールド名
'          :戻り値        ,O  ,String    ,フィールド内容
'説明      :
'履歴      :2005/06/02  KODA9
Public Function GetCodeFieldA9(ByVal SYSCLASS$, ByVal Class$, ByVal CODE$, ByVal FieldName$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String

    ''コードマスタから、指定のフィールドを引く
    sql = "SELECT " & FieldName & " from KODA9 WHERE (SYSCA9='" & SYSCLASS & "') and (SHUCA9='" & Class & "') and (CODEA9='" & CODE & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''見つからなかったら、VbNullStringを返す
        GetCodeFieldA9 = vbNullString
    Else
        ''見つかったら、指定フィールドの内容を返す
        If IsNull(rs(FieldName)) = False Then
            GetCodeFieldA9 = Trim$(rs(FieldName))
        End If
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

End Function


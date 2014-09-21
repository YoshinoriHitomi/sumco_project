Attribute VB_Name = "s_XSDCS_1_SQL"
Option Explicit
'***テーブル「XSDCS_1」へのデータアクセス関数***

Public Type typ_XSDCS_1
    CRYNUMCS1       As String       ' ブロックID
    XTALCS1         As String       ' 結晶番号
    INPOSCS1        As String       ' 結晶内位置
    HINBCS1         As String       ' 品番
    REVNUMCS1       As String       ' 品番製品番号改訂番号
    FACTORYCS1      As String       ' 品番工場
    OPECS1          As String       ' 品番操業条件
    TRANCNTFRSCS1   As String       ' 処理回数(FRS)
    CRYINDOIFRSCS1  As String       ' 状態FLG(FRS)
    CRYRESOIFRSCS1  As String       ' 実績FLG(FRS)
    RPCRYNUMCS1     As String       ' 親ブロックID
    LIVKCS1         As String       ' 生死区分
    TSTAFFCS1       As String       ' 登録社員ID
    TDAYCS1         As String       ' 登録日付
    KSTAFFCS1       As String       ' 更新社員ID
    KDAYCS1         As String       ' 更新日付
    SNDKCS1         As String       ' 送信フラグ
    SNDDAYCS1       As String       ' 送信日付
    SNDKDWHCS1      As String       ' 送信フラグ(DWH)
    SDAYDWHCS1      As String       ' 送信日付(DWH)
    SNDKSPCCS1      As String       ' 送信フラグ(SPC)
    SDAYSPCCS1      As String       ' 送信日付(SPC)
    SAKJCS1         As String       ' 削除区分
End Type

Private Const SQRT = "'"

'概要      :テーブル「XSDCS_1」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ　　:ブロックID
'           XSDCS_1　　抽出結果１配列目から（０配列目未使用）
'戻り値    :抽出の成否 Boolean
'説明      :ﾌﾞﾛｯｸIDでXSDC_1を検索
'履歴      :2011/02/28　作成 SMPK H.Ohkubo
Public Function GetXSDCS_1(sBlock As String, typXSDCS1() As typ_XSDCS_1) As Boolean
    Dim objDS       As Object
    Dim sSQL        As String
    Dim recCnt      As Integer
    Dim lRecCnt     As Long
    Dim lDtCnt      As Long
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDCS_1_SQL.bas -- Function GetXSDCS_1"
    
    '' ■該当レコード件数取得し、無い場合、仮テーブル生成
    If CheckUniqueRecordXSDCS_1(sBlock) = False Then
                
        '' 該当データ無しの場合、空の仮テーブルを生成する
        ReDim Preserve typXSDCS1(1) As typ_XSDCS_1
        
        With typXSDCS1(1)
            .CRYNUMCS1 = sBlock                     '' ブロックID
            .XTALCS1 = left(sBlock, 9) & "000"      '' 結晶番号
            .INPOSCS1 = ""                          '' 結晶内位置
            .HINBCS1 = ""                           '' 品番
            .REVNUMCS1 = ""                         '' 品番製品番号改訂番号
            .FACTORYCS1 = ""                        '' 品番工場
            .OPECS1 = ""                            '' 品番操業条件
            .TRANCNTFRSCS1 = "0"                    '' 処理回数(FRS)
            .CRYINDOIFRSCS1 = "0"                   '' 状態FLG
            .CRYRESOIFRSCS1 = "0"                   '' 実績FLG
            .RPCRYNUMCS1 = left(sBlock, 9) & "000"  '' 親ブロックID
            .LIVKCS1 = "0"                          '' 生死区分
            .TSTAFFCS1 = ""                         '' 登録社員ID
            .TDAYCS1 = ""                           '' 登録日付
            .KSTAFFCS1 = ""                         '' 更新社員ID
            .KDAYCS1 = ""                           '' 更新日付
            .SNDKCS1 = ""                           '' 送信フラグ
            .SNDDAYCS1 = ""                         '' 送信日付
            .SNDKDWHCS1 = ""                        '' 送信フラグ(DWH)
            .SDAYDWHCS1 = ""                        '' 送信日付(DWH)
            .SNDKSPCCS1 = ""                        '' 送信フラグ(SPC)
            .SDAYSPCCS1 = ""                        '' 送信日付(SPC)
            .SAKJCS1 = ""                           '' 削除区分
        End With
        
        GetXSDCS_1 = True
        Exit Function
    End If
    
    '' ■該当レコードデータ検索
    sSQL = "select"
    sSQL = sSQL & "  CRYNUMCS1"           '' ブロックID
    sSQL = sSQL & ", XTALCS1"             '' 結晶番号
    sSQL = sSQL & ", INPOSCS1"            '' 結晶内位置
    sSQL = sSQL & ", HINBCS1"             '' 品番
    sSQL = sSQL & ", REVNUMCS1"           '' 品番製品番号改訂番号
    sSQL = sSQL & ", FACTORYCS1"          '' 品番工場
    sSQL = sSQL & ", OPECS1"              '' 品番操業条件
    sSQL = sSQL & ", TRANCNTFRSCS1"       '' 処理回数(FRS)
    sSQL = sSQL & ", CRYINDOIFRSCS1"      '' 状態FLG(FRS)
    sSQL = sSQL & ", CRYRESOIFRSCS1"      '' 実績FLG(FRS)
    sSQL = sSQL & ", RPCRYNUMCS1"         '' 親ブロックID
    sSQL = sSQL & ", LIVKCS1"             '' 生死区分
    sSQL = sSQL & ", TSTAFFCS1"           '' 登録社員ID
    sSQL = sSQL & ", TDAYCS1"             '' 登録日付
    sSQL = sSQL & ", KSTAFFCS1"           '' 更新社員ID
    sSQL = sSQL & ", KDAYCS1"             '' 更新日付
    sSQL = sSQL & ", SNDKCS1"             '' 送信フラグ
    sSQL = sSQL & ", SNDDAYCS1"           '' 送信日付
    sSQL = sSQL & ", SNDKDWHCS1"          '' 送信フラグ(DWH)
    sSQL = sSQL & ", SDAYDWHCS1"          '' 送信日付(DWH)
    sSQL = sSQL & ", SNDKSPCCS1"          '' 送信フラグ(SPC)
    sSQL = sSQL & ", SDAYSPCCS1"          '' 送信日付(SPC)
    sSQL = sSQL & ", SAKJCS1"             '' 削除区分
    sSQL = sSQL & " from XSDCS_1"
    sSQL = sSQL & " where CRYNUMCS1 like '" & Trim(sBlock) & "%'"
    
    ''データを抽出する
#If SRC_200_FLG = 1 Then
    If DynSet(objDS, sSQL) = False Then
        Call MsgOut(100, sSQL, ERR_DISP_LOG)
        GetXSDCS_1 = False
        Exit Function
    End If
#Else
    If DynSet2(objDS, sSQL) = False Then
        Call MsgOut(100, sSQL, ERR_DISP_LOG)
        GetXSDCS_1 = False
        Exit Function
    End If
#End If
    
    ReDim typXSDCS1(0)
    lRecCnt = 0
    ''抽出結果を格納する
    If objDS.EOF = False Then
        Do Until objDS.EOF 'データがなくなるまで取得
            
            lRecCnt = lRecCnt + 1
            ReDim Preserve typXSDCS1(lRecCnt) As typ_XSDCS_1
            With typXSDCS1(lRecCnt)
                If IsNull(objDS.Fields("CRYNUMCS1")) = False Then .CRYNUMCS1 = objDS.Fields("CRYNUMCS1")                    '' ブロックID
                If IsNull(objDS.Fields("XTALCS1")) = False Then .XTALCS1 = objDS.Fields("XTALCS1")                          '' 結晶番号
                If IsNull(objDS.Fields("INPOSCS1")) = False Then .INPOSCS1 = objDS.Fields("INPOSCS1")                       '' 結晶内位置
                If IsNull(objDS.Fields("HINBCS1")) = False Then .HINBCS1 = objDS.Fields("HINBCS1")                          '' 品番
                If IsNull(objDS.Fields("REVNUMCS1")) = False Then .REVNUMCS1 = objDS.Fields("REVNUMCS1")                    '' 品番製品番号改訂番号
                If IsNull(objDS.Fields("FACTORYCS1")) = False Then .FACTORYCS1 = objDS.Fields("FACTORYCS1")                 '' 品番工場
                If IsNull(objDS.Fields("OPECS1")) = False Then .OPECS1 = objDS.Fields("OPECS1")                             '' 品番操業条件
                If IsNull(objDS.Fields("TRANCNTFRSCS1")) = False Then .TRANCNTFRSCS1 = objDS.Fields("TRANCNTFRSCS1")        '' 処理回数(FRS)
                If IsNull(objDS.Fields("CRYINDOIFRSCS1")) = False Then .CRYINDOIFRSCS1 = objDS.Fields("CRYINDOIFRSCS1")     '' 状態FLG
                If IsNull(objDS.Fields("CRYRESOIFRSCS1")) = False Then .CRYRESOIFRSCS1 = objDS.Fields("CRYRESOIFRSCS1")     '' 実績FLG
                If IsNull(objDS.Fields("RPCRYNUMCS1")) = False Then .RPCRYNUMCS1 = objDS.Fields("RPCRYNUMCS1")              '' 親ブロックID
                If IsNull(objDS.Fields("LIVKCS1")) = False Then .LIVKCS1 = objDS.Fields("LIVKCS1")                          '' 生死区分
                If IsNull(objDS.Fields("TSTAFFCS1")) = False Then .TSTAFFCS1 = objDS.Fields("TSTAFFCS1")                    '' 登録社員ID
                If IsNull(objDS.Fields("TDAYCS1")) = False Then .TDAYCS1 = objDS.Fields("TDAYCS1")                          '' 登録日付
                If IsNull(objDS.Fields("KSTAFFCS1")) = False Then .KSTAFFCS1 = objDS.Fields("KSTAFFCS1")                    '' 更新社員ID
                If IsNull(objDS.Fields("KDAYCS1")) = False Then .KDAYCS1 = objDS.Fields("KDAYCS1")                          '' 更新日付
                If IsNull(objDS.Fields("SNDKCS1")) = False Then .SNDKCS1 = objDS.Fields("SNDKCS1")                          '' 送信フラグ
                If IsNull(objDS.Fields("SNDDAYCS1")) = False Then .SNDDAYCS1 = objDS.Fields("SNDDAYCS1")                    '' 送信日付
                If IsNull(objDS.Fields("SNDKDWHCS1")) = False Then .SNDKDWHCS1 = objDS.Fields("SNDKDWHCS1")                 '' 送信フラグ(DWH)
                If IsNull(objDS.Fields("SDAYDWHCS1")) = False Then .SDAYDWHCS1 = objDS.Fields("SDAYDWHCS1")                 '' 送信日付(DWH)
                If IsNull(objDS.Fields("SNDKSPCCS1")) = False Then .SNDKSPCCS1 = objDS.Fields("SNDKSPCCS1")                 '' 送信フラグ(SPC)
                If IsNull(objDS.Fields("SDAYSPCCS1")) = False Then .SDAYSPCCS1 = objDS.Fields("SDAYSPCCS1")                 '' 送信日付(SPC)
                If IsNull(objDS.Fields("SAKJCS1")) = False Then .SAKJCS1 = objDS.Fields("SAKJCS1")                          '' 削除区分
            End With
            objDS.MoveNext
        Loop
    End If
    
    objDS.Close

    GetXSDCS_1 = True

proc_exit:
    '' 終了
    Exit Function

proc_err:
    Call MsgOut(100, "", ERR_DISP_LOG, "XSDCS_1")
    GetXSDCS_1 = False
    Resume proc_exit
End Function

'概要      :該当するﾚｺｰﾄﾞ有無をﾁｪｯｸ
'ﾊﾟﾗﾒｰﾀ　　:ブロックID
'      　　:戻り値       ,O  ,Boolean        　,TRUE:有/ FALSE:無（異常）
'説明      :ﾌﾞﾛｯｸIDでXSDC_1を検索
'履歴      :2011/02/28　作成 SMPK H.Ohkubo
Public Function CheckUniqueRecordXSDCS_1(sBlock As String) As Boolean
    Dim objDS       As Object
    Dim sSQL        As String
    Dim lRecCnt     As Long
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDCS_1_SQL.bas -- Function CheckUniqueRecordXSDCS_1"
    
    sSQL = "select count(*) CNT"
    sSQL = sSQL & " from XSDCS_1"
    sSQL = sSQL & " where CRYNUMCS1 like '" & Trim(sBlock) & "%'"
    
    ''データを抽出する
#If SRC_200_FLG = 1 Then
    If DynSet(objDS, sSQL) = False Then
        Call MsgOut(100, sSQL, ERR_DISP_LOG)
        CheckUniqueRecordXSDCS_1 = False
        Exit Function
    End If
#Else
    If DynSet2(objDS, sSQL) = False Then
        Call MsgOut(100, sSQL, ERR_DISP_LOG)
        CheckUniqueRecordXSDCS_1 = False
        Exit Function
    End If
#End If
    
    lRecCnt = 0

    ''抽出結果を格納する
    If objDS.EOF = False Then
        lRecCnt = objDS.Fields("CNT")
    End If
    
    objDS.Close

    If lRecCnt > 0 Then
        CheckUniqueRecordXSDCS_1 = True
    Else
        CheckUniqueRecordXSDCS_1 = False
    End If

proc_exit:
    '' 終了
    Exit Function

proc_err:
    Call MsgOut(100, "", ERR_DISP_LOG, "XSDCS_1")
    CheckUniqueRecordXSDCS_1 = False
    Resume proc_exit
End Function

'概要      :テーブル「XSDCS_1」にレコードを挿入する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型                ,説明
'      　　:pXSDCS_1　   ,I  ,typ_XSDCS_1       ,XSDCS_1更新用ﾃﾞｰﾀ
'      　　:sErrMsg　　　,O  ,String         　 ,エラーメッセージ
'      　　:戻り値       ,O  ,Boolean　,書き込みの成否
Public Function InsertXSDCS_1(pXSDCS_1 As typ_XSDCS_1) As Boolean

    Dim sSQL    As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    
    '' ■INSERT
    sSQL = ""
    sSQL = sSQL & " INSERT INTO XSDCS_1"
    sSQL = sSQL & " ("
    sSQL = sSQL & "  CRYNUMCS1"           '' ブロックID
    sSQL = sSQL & ", XTALCS1"             '' 結晶番号
    sSQL = sSQL & ", INPOSCS1"            '' 結晶内位置
    sSQL = sSQL & ", HINBCS1"             '' 品番
    sSQL = sSQL & ", REVNUMCS1"           '' 品番製品番号改訂番号
    sSQL = sSQL & ", FACTORYCS1"          '' 品番工場
    sSQL = sSQL & ", OPECS1"              '' 品番操業条件
    sSQL = sSQL & ", TRANCNTFRSCS1"       '' 処理回数(FRS)
    sSQL = sSQL & ", CRYINDOIFRSCS1"      '' 状態FLG(FRS)
    sSQL = sSQL & ", CRYRESOIFRSCS1"      '' 実績FLG(FRS)
    sSQL = sSQL & ", RPCRYNUMCS1"         '' 親ブロックID
    sSQL = sSQL & ", LIVKCS1"             '' 生死区分
    sSQL = sSQL & ", TSTAFFCS1"           '' 登録社員ID
    sSQL = sSQL & ", TDAYCS1"             '' 登録日付
    sSQL = sSQL & " )VALUES"
    sSQL = sSQL & " ("
    
    With pXSDCS_1
        sSQL = sSQL & " " & Cnv2String2(.CRYNUMCS1) & ""                        '' ブロックID
        sSQL = sSQL & "," & Cnv2String2(.XTALCS1) & ""                          '' 結晶番号
        sSQL = sSQL & "," & Cnv2Number(.INPOSCS1) & ""                          '' 結晶内位置
        sSQL = sSQL & "," & Cnv2String2(.HINBCS1) & ""                          '' 品番
        sSQL = sSQL & "," & Cnv2Number(.REVNUMCS1) & ""                         '' 品番製品番号改訂番号
        sSQL = sSQL & "," & Cnv2String2(.FACTORYCS1) & ""                       '' 品番工場
        sSQL = sSQL & "," & Cnv2String2(.OPECS1) & ""                           '' 品番操業条件
        sSQL = sSQL & "," & Cnv2Number(.TRANCNTFRSCS1) & ""                     '' 処理回数(FRS)
        sSQL = sSQL & "," & Cnv2String2(.CRYINDOIFRSCS1) & ""                   '' 状態FLG(FRS)
        sSQL = sSQL & "," & Cnv2String2(.CRYRESOIFRSCS1) & ""                   '' 実績FLG(FRS)
        sSQL = sSQL & "," & Cnv2String2(.RPCRYNUMCS1) & ""                      '' 親ブロックID
        sSQL = sSQL & "," & Cnv2String2(.LIVKCS1) & ""                          '' 生死区分
        sSQL = sSQL & "," & Cnv2String2(.TSTAFFCS1) & ""                        '' 登録社員ID
        sSQL = sSQL & ",to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')"   '' 登録日付
        sSQL = sSQL & " )"
    End With
    
'    Debug.Print sSql
    
#If SRC_200_FLG = 1 Then
    If SqlExec(sSQL) <> 1 Then
        Call MsgOut(100, sSQL, ERR_DISP_LOG, "XSDCS_1")
        InsertXSDCS_1 = False
        Exit Function
    End If
#Else
    If SqlExec2(sSQL) <> 1 Then
        Call MsgOut(0, "DB登録失敗" & sSQL, ERR_DISP_LOG, "CRYNUMCS1")
        InsertXSDCS_1 = False
        Exit Function
    End If
#End If
    
    InsertXSDCS_1 = True

proc_exit:
    '' 終了
    Exit Function

proc_err:
    Call MsgOut(0, "DB登録失敗" & sSQL, ERR_DISP_LOG, "CRYNUMCS1")
    InsertXSDCS_1 = False
    Resume proc_exit

End Function

'●更新項目を構造体にセットして引き渡す

'概要      :テーブル「XSDCS_1」を更新する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型              ,説明
'          :sBlock        ,I   ,String          ,ブロックID
'          :records()     ,O   ,typ_XSDCS_1     ,更新レコード
'          :戻り値        ,O   ,Boolean         ,抽出の成否
'説明      :
Public Function UpdateXSDCS_1(sBlock As String, records As typ_XSDCS_1) As Boolean
    Dim sql     As String       'SQL全体
    
    On Error GoTo proc_err

    With records
        sql = ""
        sql = sql & "UPDATE XSDCS_1 "
        sql = sql & "SET "
        
        If .XTALCS1 <> "" And left(.XTALCS1, 1) <> vbNullChar Then sql = sql & "  XTALCS1 = " & Cnv2String(.XTALCS1)                                ' 結晶番号
        If .INPOSCS1 <> "" And left(.INPOSCS1, 1) <> vbNullChar Then sql = sql & ", INPOSCS1 = " & Cnv2Number(.INPOSCS1)                            ' 結晶内位置
        If .HINBCS1 <> "" And left(.HINBCS1, 1) <> vbNullChar Then sql = sql & ", HINBCS1 = " & Cnv2String(.HINBCS1)                                ' 品番
        If .REVNUMCS1 <> "" And left(.REVNUMCS1, 1) <> vbNullChar Then sql = sql & ", REVNUMCS1 = " & Cnv2Number(.REVNUMCS1)                        ' 品番製品番号改訂番号
        If .FACTORYCS1 <> "" And left(.FACTORYCS1, 1) <> vbNullChar Then sql = sql & ", FACTORYCS1 = " & Cnv2String(.FACTORYCS1)                    ' 品番工場
        If .OPECS1 <> "" And left(.OPECS1, 1) <> vbNullChar Then sql = sql & ", OPECS1 = " & Cnv2String(.OPECS1)                                    ' 品番操業条件
        If .TRANCNTFRSCS1 <> "" And left(.TRANCNTFRSCS1, 1) <> vbNullChar Then sql = sql & ", TRANCNTFRSCS1 = " & Cnv2Number(.TRANCNTFRSCS1)        ' 処理回数(FRS)
        If .CRYINDOIFRSCS1 <> "" And left(.CRYINDOIFRSCS1, 1) <> vbNullChar Then sql = sql & ", CRYINDOIFRSCS1 = " & Cnv2String(.CRYINDOIFRSCS1)    ' 状態FLG(FRS)
        If .CRYRESOIFRSCS1 <> "" And left(.CRYRESOIFRSCS1, 1) <> vbNullChar Then sql = sql & ", CRYRESOIFRSCS1 = " & Cnv2String(.CRYRESOIFRSCS1)    ' 実績FLG(FRS)
        
        If .RPCRYNUMCS1 <> sBlock Then
            If .RPCRYNUMCS1 <> "" And left(.RPCRYNUMCS1, 1) <> vbNullChar Then sql = sql & ", RPCRYNUMCS1 = " & Cnv2String(.RPCRYNUMCS1)            ' 親ブロックID
        End If
        
        If .LIVKCS1 <> "" And left(.LIVKCS1, 1) <> vbNullChar Then sql = sql & ", LIVKCS1 = " & Cnv2String(.LIVKCS1)                                ' 生死区分
        If .KSTAFFCS1 <> "" And left(.KSTAFFCS1, 1) <> vbNullChar Then sql = sql & ", KSTAFFCS1 = " & Cnv2String(.KSTAFFCS1)                        ' 更新社員ID
        sql = sql & ", KDAYCS1 = to_date(" & Cnv2String(gsSysdate) & ",'yyyy/mm/dd hh24:mi:ss')"                                                    ' 更新日付
        
        sql = sql & " where CRYNUMCS1 = '" & sBlock & "'"
    
    End With
'    Debug.Print sql

#If SRC_200_FLG = 1 Then
    If SqlExec(sql) < 0 Then
        Call MsgOut(100, sql, ERR_DISP_LOG, "XSDCS_1")
        UpdateXSDCS_1 = False
        Exit Function
    End If
#Else
    If SqlExec2(sql) < 0 Then
        Call MsgOut(100, sql, ERR_DISP_LOG, "XSDCS_1")
        UpdateXSDCS_1 = False
        Exit Function
    End If
#End If

    UpdateXSDCS_1 = True

proc_exit:
    '' 終了
    Exit Function

proc_err:
    Call MsgOut(0, "DB更新失敗", ERR_DISP_LOG, "CRYNUMCS1")
    UpdateXSDCS_1 = False
    Resume proc_exit

End Function

'概要      :テーブル「XSDCS_1」を死ロットにする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型              ,説明
'          :sBlock        ,I   ,String          ,ブロックID
'          :sStaff        ,I   ,String          ,担当者ID
'          :戻り値        ,O   ,Boolean         ,抽出の成否
'説明      :
Public Function UpdateXSDCS_1Delete(sBlock As String, sStaff As String) As Boolean
    Dim sql         As String       'SQL全体
    
    On Error GoTo proc_err

    '' ■該当レコード件数取得し、無い場合、更新処理終了
    If CheckUniqueRecordXSDCS_1(sBlock) = False Then
        '' 更新処理終了
        UpdateXSDCS_1Delete = True
        Exit Function
    End If
    
    '' ■UPDATE
    sql = ""
    sql = sql & "UPDATE XSDCS_1 "
    sql = sql & "SET "
    
    sql = sql & "  LIVKCS1 = '1'"                                                               ' 生死区分
    sql = sql & ", KSTAFFCS1 = " & Cnv2String(sStaff)                                           ' 更新社員ID
    sql = sql & ", KDAYCS1 = to_date(" & Cnv2String(gsSysdate) & ",'yyyy/mm/dd hh24:mi:ss')"    ' 更新日付
    sql = sql & " where CRYNUMCS1 = '" & sBlock & "'"

'    Debug.Print sql

#If SRC_200_FLG = 1 Then
    If SqlExec(sql) < 0 Then
        Call MsgOut(100, sql, ERR_DISP_LOG, "XSDCS_1")
        UpdateXSDCS_1Delete = False
        Exit Function
    End If
#Else
    If SqlExec2(sql) < 0 Then
        Call MsgOut(100, sql, ERR_DISP_LOG, "XSDCS_1")
        UpdateXSDCS_1Delete = False
        Exit Function
    End If
#End If

    UpdateXSDCS_1Delete = True

proc_exit:
    '' 終了
    Exit Function

proc_err:
    Call MsgOut(0, "DB更新失敗", ERR_DISP_LOG, "CRYNUMCS1")
    UpdateXSDCS_1Delete = False
    Resume proc_exit

End Function

'概要      :テーブル「XSDCS_1」登録・更新
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型             ,説明
'          :sBlockID      ,I    ,String         ,ブロックID
'          :sCut          ,I    ,String         ,CUT位置
'          :sHin          ,I    ,String         ,品番
'          :sRev          ,I    ,String         ,リビジョン
'          :sTranCnt      ,I    ,String         ,処理回数
'          :sFRSKbn       ,I    ,String         ,FRS状態FLG
'          :sResult       ,I    ,String         ,FRS実績FLG
'          :sStaff        ,I    ,String         ,担当者ID
'          :sXtal         ,I    ,String         ,結晶番号
'          :sRpBlockId    ,I    ,String         ,親ブロックID
'          :戻り値        ,O    ,Boolean        ,実行の成否
'説明      :
Public Function CreateOrUpdateXSDCS_1(sBlockId As String _
                                    , sCut As String _
                                    , sHin As String _
                                    , sRev As String _
                                    , sTranCnt As String _
                                    , sFRSKbn As String _
                                    , sResult As String _
                                    , sStaff As String _
                                    , sXtal As String _
                                    , sRpBlockId As String _
                                    ) As Boolean
    Dim tXSDCS_1        As typ_XSDCS_1
    
    CreateOrUpdateXSDCS_1 = False
    
    Call LogInit
    Call GetSysdate
    
    tXSDCS_1.CRYNUMCS1 = sBlockId                           '' ブロックID
    tXSDCS_1.XTALCS1 = sXtal                                '' 結晶番号
    tXSDCS_1.INPOSCS1 = sCut                                '' 結晶内位置
    tXSDCS_1.HINBCS1 = sHin                                 '' 品番
    tXSDCS_1.REVNUMCS1 = left(sRev, 2)                      '' 品番製品番号改訂番号
    tXSDCS_1.FACTORYCS1 = Mid(sRev, 3, 1)                   '' 品番工場
    tXSDCS_1.OPECS1 = Right(sRev, 1)                        '' 品番操業条件
    tXSDCS_1.TRANCNTFRSCS1 = sTranCnt                       '' 処理回数(FRS)
    tXSDCS_1.CRYINDOIFRSCS1 = sFRSKbn                       '' 状態FLG(FRS)
    tXSDCS_1.CRYRESOIFRSCS1 = sResult                       '' 実績FLG(FRS)
    tXSDCS_1.RPCRYNUMCS1 = sRpBlockId                       '' 親ブロックID
    tXSDCS_1.LIVKCS1 = "0"                                  '' 生死区分
    tXSDCS_1.TSTAFFCS1 = sStaff                             '' 登録社員ID
    
    '' 対象ブロックIDデータ存在チェック
    If CheckUniqueRecordXSDCS_1(sBlockId) = False Then
        
        '' 対象ブロックIDのXSDCS_1が存在しない場合、追加（切断有り）
        
        '' ■DB登録
        If InsertXSDCS_1(tXSDCS_1) = False Then
            Exit Function
        End If
        
'' Chg Start 2011/05/10 SMPK H.Ohkubo
        '' ■親ブロックIDのXSDCS_1.LIVKCS1="1"
''        If UpdateXSDCS_1Delete(sRpBlockId, sStaff) = False Then
''            Exit Function
''        End If
        If sBlockId <> sRpBlockId Then
            If UpdateXSDCS_1Delete(sRpBlockId, sStaff) = False Then
                Exit Function
            End If
        End If
'' Chg End 2011/05/10 SMPK H.Ohkubo

    Else
        '' 対象ブロックIDのXSDCS_1が存在する場合、更新（切断無し）
        
        '' ■DB更新
        If UpdateXSDCS_1(sBlockId, tXSDCS_1) = False Then
            Exit Function
        End If
        
    End If
    
    CreateOrUpdateXSDCS_1 = True

End Function

' @(f)
' 機能      : SQL文字列変換関数
'
' 返り値    : '<入力文字列>' or NULL
'
' 引き数    : 変換対象文字列
'
' 機能説明  : 文字列がnullであれば"NULL"を、そうでなければシングルコーテーションをつけて出力する
Private Function Cnv2String(vinput) As String
    If IsNull(vinput) Or vinput = "NULL" Then
        vinput = ""
    End If
    
    If vinput = "" Then
        Cnv2String = "NULL"
    Else
        Cnv2String = SQRT & vinput & SQRT
    End If
    
End Function

' @(f)
' 機能      : SQL文字列変換関数
'
' 返り値    : '<入力文字列>' or NULL
'
' 引き数    : 変換対象文字列
'
' 機能説明  : 文字列がnullであれば"NULL"を、そうでなければシングルコーテーションをつけて出力する
Private Function Cnv2String2(vinput) As String
    If IsNull(vinput) Or vinput = "NULL" Then
        vinput = ""
    End If
    
    If vinput = "" Then
        Cnv2String2 = SQRT & " " & SQRT
    Else
        Cnv2String2 = SQRT & vinput & SQRT
    End If
    
End Function

' @(f)
' 機能      : SQL数値変換関数
'
' 返り値    : <入力数値> or NULL
'
' 引き数    : 変換対象数値
'
' 機能説明  : 渡された数値がNULLであれば"NULL"をそうでなければそのまま出力する
Private Function Cnv2Number(vinput) As String
    If IsNull(vinput) Or vinput = "NULL" Then
        vinput = ""
    End If
    
    If vinput = "" Then
        Cnv2Number = "NULL"
    Else
        Cnv2Number = vinput
    End If
End Function

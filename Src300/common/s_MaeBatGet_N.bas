Attribute VB_Name = "s_MaeBatGet"

'=================================================================================
'概要      :前バッチ結晶長取得処理
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型              , 説明
'          :CRYNUM          , I  ,String           , 結晶番号
'      　　:戻り値          , O  ,Double　         , 前バッチ結晶長
'説明      :
'履歴      :2011/09/27 Marushita
Public Function GetMaeBatLen(ByVal CRYNUM As String, Optional ByVal iKata As Integer = 0) As Double
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    Dim dBatLen     As Double               'バッチ結晶長
    Dim p_CRYNUM    As String               '前結晶番号取得用
    Dim p_wgtTop    As Double               '前引上情報取得用(前Top取量)
    Dim p_DM        As Double               '前引上情報取得用(前直径平均)
    Dim p_wgtTA     As Double               '前引上情報取得用(前テイル重量)
    Dim p_LENTK     As Long                 '前引上情報取得用(前引上長)
    Dim p_wgtKata   As Double               '前引上情報取得用(前肩重量)
    Dim n_wgtTop    As Double               '現在引上情報取得用(Top取量)
    Dim n_DM        As Double               '現在引上情報取得用(直径平均)
    Dim n_wgtTA     As Double               '現在引上情報取得用(テイル重量)
    Dim n_LENTK     As Long                 '現在引上情報取得用(引上長)
    Dim n_wgtKata   As Double               '現在引上情報取得用(肩重量)
    Dim iMaeBatFlg  As Integer              '結晶長計算FLG
    Dim N_CRYNUM    As String               '現在結晶番号セット用
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_MaeBatLen.bas -- Function GetMaeBatLen"
    
    GetMaeBatLen = 0
        
    dBatLen = 0
    
    'CRYNUMの9桁目が"A"のとき処理を抜ける
    If Mid(CRYNUM, 9, 1) = "A" Then
        Exit Function
    Else
        '現在引上の結晶情報を取得(Top取量、直径平均、引上長、テイル重量、肩重量)
        If GetXSDC1Info(CRYNUM, n_wgtTop, n_DM, n_LENTK, n_wgtTA, n_wgtKata) = FUNCTION_RETURN_FAILURE Then
            Exit Function
        End If
    End If
    
    '現在結晶番号の初期セット
    N_CRYNUM = CRYNUM
        
    Do While iMaeBatFlg = 0
        '前引上の結晶番号を取得(現在結晶番号)
        If GetPreCrynum(N_CRYNUM, p_CRYNUM) = FUNCTION_RETURN_FAILURE Then
            iMaeBatFlg = 1
        Else
            '前引上の結晶情報を取得(前Top取量、前直径平均、前引上長、前テイル重量)
            If GetXSDC1Info(p_CRYNUM, p_wgtTop, p_DM, p_LENTK, p_wgtTA, p_wgtKata) = FUNCTION_RETURN_FAILURE Then
                iMaeBatFlg = 1
            Else
                '前バッチ結晶長の計算
                '肩重量込み
                If iKata = 0 Then
                    '前バッチ結晶長 = 前Top取量 / (前断面積 * 0.00233) + 前引上長 + ((前テイル重量 + 肩重量) / (断面積 * 0.00233))
                    dBatLen = dBatLen + (p_wgtTop / (AreaOfCircle(p_DM) * HIJU_SILICONE)) + p_LENTK + ((p_wgtTA + n_wgtKata) / (AreaOfCircle(n_DM) * HIJU_SILICONE))
                Else
                    '前バッチ結晶長 = 前Top取量 / (前断面積 * 0.00233) + 前引上長 + ((前テイル重量 + 肩重量) / (断面積 * 0.00233))
                    dBatLen = dBatLen + (p_wgtTop / (AreaOfCircle(p_DM) * HIJU_SILICONE)) + p_LENTK + (p_wgtTA) / (AreaOfCircle(n_DM) * HIJU_SILICONE)
                    iKata = 0
                End If
                '現在結晶番号のセット
                N_CRYNUM = p_CRYNUM
                '前肩重量、前断面積を現在にセット
                n_wgtKata = p_wgtKata
                n_DM = p_DM
            End If
        End If
    Loop
    
    GetMaeBatLen = dBatLen
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    GetMaeBatLen = 0
    gErr.HandleError
    Resume proc_exit
End Function

'=================================================================================
'概要      :前バッチ残液取得処理
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型              , 説明
'          :CRYNUM          , I  ,String           , 結晶番号
'      　　:戻り値          , O  ,Long  　         , 前バッチ残液
'説明      :
'履歴      :2011/09/27 Marushita
Public Function GetMaeBatZan(ByVal CRYNUM As String) As Double
    
    Dim sqlWhere As String
    
    Dim lZaneki     As Double               '残液量
    Dim p_CRYNUM    As String               '前結晶番号取得用
    Dim p_wgtTop    As Double               '前引上情報取得用(前BtTop取量)
    Dim p_wgtTA     As Double               '前引上情報取得用(前Bt仕込量)
    Dim p_wgWGHTTK  As Long                 '前引上情報取得用(前Bt炉上重量)
    Dim iMaeBatFlg  As Integer              '結晶長計算FLG
    Dim N_CRYNUM    As String               '現在結晶番号セット用
    Dim tblXSDC1()  As typ_XSDC1            'XSDC1データ取得用
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_MaeBatGet.bas -- Function GetMaeBatZan"
    
    GetMaeBatZan = 0
        
    lZaneki = 0
    
    '現在結晶番号の初期セット
    N_CRYNUM = CRYNUM
        
    Do While iMaeBatFlg = 0
        '前引上連番の結晶番号を取得(現在結晶番号)
        If GetPreRenCrynum(N_CRYNUM, p_CRYNUM) = FUNCTION_RETURN_FAILURE Then
            iMaeBatFlg = 1
        Else
            'WHERE条件
            sqlWhere = "WHERE XTALC1 = '" & p_CRYNUM & "'"
            'レコードセットの取得(失敗したらプロシージャから抜ける）
            If DBDRV_GetXSDC1(tblXSDC1, sqlWhere) = FUNCTION_RETURN_FAILURE Then
                iMaeBatFlg = 1
            Else
                '前残液の計算
                '前残液 = 前Bt前残液 + 前Bt仕込量 - 前Bt炉上重量 - 前BtTop取量
                lZaneki = lZaneki + tblXSDC1(1).PUCHAGC1 - tblXSDC1(1).PUWC1 - tblXSDC1(1).PUTCUTWC1
                'lZaneki = lZaneki + tblXSDC1(1).PUCHAGC1 - tblXSDC1(1).PUWC1 - tblXSDC1(1).WGHTTOC1
                '現在結晶番号のセット
                N_CRYNUM = p_CRYNUM
            End If
        End If
    Loop
    
    GetMaeBatZan = lZaneki
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    GetMaeBatZan = 0
    gErr.HandleError
    Resume proc_exit
End Function

'=================================================================================
'概要       :前引上の結晶番号を取得する。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :sCrynum        ,I  ,String             　      ,現在結晶番号
'           :sP_Crynum      ,O  ,String             　      ,前引上の結晶番号
'           :戻り値         ,O  ,FUNCTION_RETURN            ,
'説明       :
'履歴       :2011/09/27 Marushita
Public Function GetPreCrynum(ByVal sCrynum As String, ByRef sP_Crynum As String) As FUNCTION_RETURN
    
    On Error GoTo Err
    GetPreCrynum = FUNCTION_RETURN_FAILURE
                
    Dim sCrynum9 As String
    Dim iCrynum9 As Integer
            
    sP_Crynum = ""
    
    'sCrynumの9桁目をセット
    sCrynum9 = Mid(sCrynum, 9, 1)
    
    '"A","1"はエラーで返す
    If sCrynum9 = "A" Or sCrynum9 = "1" Then
        Exit Function
    Else
        iCrynum9 = Asc(sCrynum9)
        sP_Crynum = Mid(sCrynum, 1, 8) & Chr(iCrynum9 - 1) & Mid(sCrynum, 10, 3)
    End If
    
    GetPreCrynum = FUNCTION_RETURN_SUCCESS
    
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    
End Function

'=================================================================================
'概要       :指定結晶番号の結晶情報を取得する。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :sCRYNUM        ,I  ,String    ,前引上結晶番号
'           :dWgtTop        ,O  ,Double    ,トップ重量実績値
'           :dDM            ,O  ,Double    ,直径１～３の平均
'           :lLentk         ,O  ,Long      ,引上長
'           :dwgtTA         ,O  ,Double    ,テイル重量実績値
'           :dwgtKata　     ,O  ,Double    ,肩重量
'           :戻り値         ,O  ,FUNCTION_RETURN
'説明       :
'履歴       :2011/10/19 Marushita　関数名・肩重量項目の追加
Public Function GetXSDC1Info(ByVal sCrynum As String, _
                            ByRef dwgtTop As Double, _
                            ByRef dDM As Double, _
                            ByRef lLenTK As Long, _
                            ByRef dwgtTA As Double, _
                            ByRef dwgtKata As Double) As FUNCTION_RETURN

    On Error GoTo proc_err
    
    Dim sql As String
    Dim rs As OraDynaset
        
    GetXSDC1Info = FUNCTION_RETURN_FAILURE
    
    sql = " SELECT WGHTTOC1, LENTKC1, WGHTTAC1, PUTCUTWC1, "
    sql = sql & " (DIA1C1 + DIA2C1 + DIA3C1) / 3.0 AS DM "
    sql = sql & " FROM XSDC1 "
    sql = sql & " WHERE XTALC1 = '" & sCrynum & "'"
        
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        dwgtTop = rs("PUTCUTWC1")           ''重量（TOP）PUTCUTWC1に変更　2011/11/10
        'dwgtTop = rs("WGHTTOC1")           ''重量（TOP）
        dDM = rs("DM")                     ''直胴直径(平均値)
        lLenTK = rs("LENTKC1")             ''引上長
        dwgtTA = rs("WGHTTAC1")            ''重量（TAIL）
        dwgtKata = rs("WGHTTOC1")         ''肩重量 WGHTTOC1に変更　2011/11/10
        'dwgtKata = rs("PUTCUTWC1")         ''肩重量
    End If
    rs.Close
    
    GetXSDC1Info = FUNCTION_RETURN_SUCCESS
    
    On Error GoTo 0
    Exit Function

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    GetXSDC1Info = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
    
End Function

'=================================================================================
'概要       :前引上連番の結晶番号を取得する。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :sCrynum        ,I  ,String             　      ,現在結晶番号
'           :sP_Crynum      ,O  ,String             　      ,前引上の結晶番号
'           :戻り値         ,O  ,FUNCTION_RETURN            ,
'説明       :
'履歴       :2011/10/14 Marushita
Public Function GetPreRenCrynum(ByVal sCrynum As String, ByRef sP_Crynum As String) As FUNCTION_RETURN
    
    On Error GoTo Err
    GetPreRenCrynum = FUNCTION_RETURN_FAILURE
                
    Dim sql As String
    Dim rs As OraDynaset
    
    Dim sUpGrp      As String               '引上グループ番号
    Dim iRenban     As Integer              'グループ内連番
            
    sUpGrp = ""
    sP_Crynum = ""
    iRenban = 0
    
    sql = " SELECT RENBAN,GROUPUPINDNO "
    sql = sql & " FROM XSDC1,TBCMH001 "
    sql = sql & " WHERE XTALC1 = '" & sCrynum & "'"
    sql = sql & " AND HISIJIC1 = UPINDNO "
    sql = sql & " AND LIVK <> '1' "
        
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        'グループ番号・グループ内連番を取得
        sUpGrp = rs("GROUPUPINDNO")      ''グループ番号
        iRenban = rs("RENBAN")           ''連番
    End If
    rs.Close
        
    '連番が１以下はエラーで返す
    If iRenban <= 1 Then
        Exit Function
    Else
        iRenban = iRenban - 1
    End If
    
    sql = " SELECT XTALC1 "
    sql = sql & " FROM XSDC1,TBCMH001 "
    sql = sql & " WHERE GROUPUPINDNO = '" & sUpGrp & "'"
    sql = sql & " AND RENBAN = " & iRenban & " "
    sql = sql & " AND LIVK <> '1' "
    sql = sql & " AND HISIJIC1 = UPINDNO "
        
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        '結晶番号を取得
        sP_Crynum = rs("XTALC1")           ''結晶番号
    End If
    rs.Close
    
    GetPreRenCrynum = FUNCTION_RETURN_SUCCESS
    
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    
End Function


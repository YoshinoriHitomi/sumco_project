Attribute VB_Name = "s_cmzcF_cmic001SXL"
Option Explicit

' 結晶サンプル仕様
Public Type typ_SpSXLSamp
    hin As tFullHinban      ' 品番

    HSXRHWYS As String * 1  ' 処理方法(Rs)
    HSXRSPOT As String * 1  ' 測定点数(Rs) -> Heavy

    HSXONHWS As String * 1  ' 処理方法(Oi)
    HSXONKWY As String * 2  ' 検査方法(Oi)
    HSXONSPH As String * 1  ' 測定方法(Oi)
    HSXONSPT As String * 1  ' 測定点数(Oi) -> Heavy
    HSXONSPI As String * 1  ' 測定位置(Oi)

    HSXBM1HS As String * 1  ' 処理方法(B1)
    HSXBM1SH As String * 1  ' 測定方法(B1)
    HSXBM1ST As String * 1  ' 測定点数(B1)
    HSXBM1SR As String * 1  ' 除外領域(B1)
    HSXBM1NS As String * 2  ' 熱処理法(B1)
    HSXBM1SZ As String * 1  ' 測定条件(B1)
    HSXBM1ET As Integer     ' 選択エッチ(B1)

    HSXBM2HS As String * 1  ' 処理方法(B2)
    HSXBM2SH As String * 1  ' 測定方法(B2)
    HSXBM2ST As String * 1  ' 測定点数(B2)
    HSXBM2SR As String * 1  ' 除外領域(B2)
    HSXBM2NS As String * 2  ' 熱処理法(B2)
    HSXBM2SZ As String * 1  ' 測定条件(B2)
    HSXBM2ET As Integer     ' 選択エッチ(B2)

    HSXBM3HS As String * 1  ' 処理方法(B3)
    HSXBM3SH As String * 1  ' 測定方法(B3)
    HSXBM3ST As String * 1  ' 測定点数(B3)
    HSXBM3SR As String * 1  ' 除外領域(B3)
    HSXBM3NS As String * 2  ' 熱処理法(B3)
    HSXBM3SZ As String * 1  ' 測定条件(B3)
    HSXBM3ET As Integer     ' 選択エッチ(B3)

    HSXOF1HS As String * 1  ' 処理方法(L1)
    HSXOF1SH As String * 1  ' 測定方法(L1)
    HSXOF1ST As String * 1  ' 測定点数(L1)
    HSXOF1SR As String * 1  ' 除外領域(L1)
    HSXOF1NS As String * 2  ' 熱処理法(L1)
    HSXOF1SZ As String * 1  ' 測定条件(L1)
    HSXOF1ET As Integer     ' 選択エッチ(L1)

    HSXOF2HS As String * 1  ' 処理方法(L2)
    HSXOF2SH As String * 1  ' 測定方法(L2)
    HSXOF2ST As String * 1  ' 測定点数(L2)
    HSXOF2SR As String * 1  ' 除外領域(L2)
    HSXOF2NS As String * 2  ' 熱処理法(L2)
    HSXOF2SZ As String * 1  ' 測定条件(L2)
    HSXOF2ET As Integer     ' 選択エッチ(L2)

    HSXOF3HS As String * 1  ' 処理方法(L3)
    HSXOF3SH As String * 1  ' 測定方法(L3)
    HSXOF3ST As String * 1  ' 測定点数(L3)
    HSXOF3SR As String * 1  ' 除外領域(L3)
    HSXOF3NS As String * 2  ' 熱処理法(L3)
    HSXOF3SZ As String * 1  ' 測定条件(L3)
    HSXOF3ET As Integer     ' 選択エッチ(L3)

    HSXOF4HS As String * 1  ' 処理方法(L4)
    HSXOF4SH As String * 1  ' 測定方法(L4)
    HSXOF4ST As String * 1  ' 測定点数(L4)
    HSXOF4SR As String * 1  ' 除外領域(L4)
    HSXOF4NS As String * 2  ' 熱処理法(L4)
    HSXOF4SZ As String * 1  ' 測定条件(L4)
    HSXOF4ET As Integer     ' 選択エッチ(L4)

    HSXCNHWS As String * 1  ' 処理方法(Cs)
    CS_FROMTO As Boolean    ' CsはFromTo保証である

    HSXDENHS As String * 1  ' 処理方法(GD/DEN)
    HSXLDLHS As String * 1  ' 処理方法(GD/LDL)
    HSXDVDHS As String * 1  ' 処理方法(GD/DVD2)

    HSXLTHWS As String * 1  ' 処理方法(T)
    HSXLTSPI As String * 1  ' 測定位置(T)

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP As String * 1  ' DK温度
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    'Add Start 2010/12/10 SMPK Miyata
    HSXCHS   As String * 1  ' 処理方法(C)
    HSXCSZ   As String * 1  ' 測定条件(C)
    HSXCJHS  As String * 1  ' 処理方法(CJ)
    HSXCJNS  As String * 2  ' 熱処理法(CJ)
    HSXCJLTHS As String * 1 ' 処理方法(CJ LT)
    HSXCJLTNS As String * 2 ' 熱処理法(CJ LT)
    HSXCJ2HS As String * 1  ' 処理方法(CJ2)
    HSXCJ2NS As String * 2  ' 熱処理法(CJ2)
    'Add End   2010/12/10 SMPK Miyata
End Type

' 結晶サンプルテーブル
Public Type typ_SXLSample
    CRYINDRS As String * 1  ' 検査項目(Rs)
    CRYINDOI As String * 1  ' 検査項目(Oi)
    CRYINDB1 As String * 1  ' 検査項目(B1)
    CRYINDB2 As String * 1  ' 検査項目(B2）
    CRYINDB3 As String * 1  ' 検査項目(B3)
    CRYINDL1 As String * 1  ' 検査項目(L1)
    CRYINDL2 As String * 1  ' 検査項目(L2)
    CRYINDL3 As String * 1  ' 検査項目(L3)
    CRYINDL4 As String * 1  ' 検査項目(L4)
    CRYINDCS As String * 1  ' 検査項目(Cs)
    CRYINDGD As String * 1  ' 検査項目(GD)
    CRYINDT As String * 1   ' 検査項目(T)
    CRYINDEP As String * 1  ' 検査項目(EPD)
    CRYINDX As String * 1   ' 検査項目(X)       '2009/07/24追加 SETsw kubota
    'Add Start 2010/12/10 SMPK Miyata
    CRYINDC     As String * 1   ' 検査項目(C)
    CRYINDCJ    As String * 1   ' 検査項目(CJ)
    CRYINDCJLT  As String * 1   ' 検査項目(CJ LT)
    CRYINDCJ2   As String * 1   ' 検査項目(CJ2)
    'Add End   2010/12/10 SMPK Miyata
End Type

'概要      :製品仕様SXLデータの取得ドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名          ,IO ,型               ,説明
'      　　:pSpSXLSamp　　　,O  ,typ_SpSXLSamp  　,結晶サンプル仕様
'      　　:戻り値          ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function scmzc_getSXL(pSpSXLSamp As typ_SpSXLSamp) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmic001SXL.bas -- Function scmzc_getSXL"

    '' 製品仕様の取得
    'C-OSF3ﾌﾗｸﾞ追加　07/06/11 ooba
    'DK温度追加      08/08/25 Systech
    'Cs検査頻度_位 追加 09/01/06 ooba
    'Cu-deco(C,CJ,CJLT,CJ2)追加　Add 2010/12/10 SMPK Miyata
    sql = "select " & _
          "E018HSXRHWYS, E018HSXRSPOT, E019HSXONHWS, E019HSXONKWY, E019HSXONSPH, " & _
          "E019HSXONSPT, E019HSXONSPI, E019HSXCNHWS, E019HSXLTHWS, E020HSXBM1HS, " & _
          "case when E019HSXCNHWS in ('H','S') and (E019HSXCNMIN>0) and (E019HSXCNMAX>0) then 1 else 0 end as CS_FROMTO, " & _
          "E019HSXLTHWS, E019HSXLTSPI, E019HSXCNKHI, E020HSXBM1HS, " & _
          "E020HSXBM1SH, E020HSXBM1ST, E020HSXBM1SR, E020HSXBM1NS, E020HSXBM1SZ, " & _
          "E020HSXBM1ET, E020HSXBM2HS, E020HSXBM2SH, E020HSXBM2ST, E020HSXBM2SR, " & _
          "E020HSXBM2NS, E020HSXBM2SZ, E020HSXBM2ET, E020HSXBM3HS, E020HSXBM3SH, " & _
          "E020HSXBM3ST, E020HSXBM3SR, E020HSXBM3NS, E020HSXBM3SZ, E020HSXBM3ET, " & _
          "E020HSXOF1HS, E020HSXOF1SH, E020HSXOF1ST, E020HSXOF1SR, E020HSXOF1NS, " & _
          "E020HSXOF1SZ, E020HSXOF1ET, E020HSXOF2HS, E020HSXOF2SH, E020HSXOF2ST, " & _
          "E020HSXOF2SR, E020HSXOF2NS, E020HSXOF2SZ, E020HSXOF2ET, E020HSXOF3HS, " & _
          "E020HSXOF3SH, E020HSXOF3ST, E020HSXOF3SR, E020HSXOF3NS, E020HSXOF3SZ, " & _
          "E020HSXOF3ET, E020HSXOF4HS, E020HSXOF4SH, E020HSXOF4ST, E020HSXOF4SR, " & _
          "E020HSXOF4NS, E020HSXOF4SZ, E020HSXOF4ET, E020HSXDENHS, E020HSXDVDHS, E020HSXLDLHS, " & _
          "E036.COSF3FLAG," & _
          "E020.HSXCHS,    E020.HSXCSZ,   E020.HSXCJHS,  E020.HSXCJNS, " & _
          "E020.HSXCJLTHS, E020.HSXCJLTNS,E020.HSXCJ2HS, E020.HSXCJ2NS, " & _
          " NVL(E036.HSXDKTMP, ' ') as HSXDKTMP" & _
          " from VECME001, TBCME036 E036, TBCME020 E020" & _
          " where E018HINBAN=E036.HINBAN and E018MNOREVNO=E036.MNOREVNO and E018FACTORY=E036.FACTORY and E018OPECOND=E036.OPECOND" & _
          " and E018HINBAN=E020.HINBAN and E018MNOREVNO=E020.MNOREVNO and E018FACTORY=E020.FACTORY and E018OPECOND=E020.OPECOND" & _
          " and E018HINBAN='" & pSpSXLSamp.hin.hinban & "' and E018MNOREVNO=" & pSpSXLSamp.hin.mnorevno & _
          " and E018FACTORY='" & pSpSXLSamp.hin.Factory & "' and E018OPECOND='" & pSpSXLSamp.hin.OpeCond & "'"
    'Add Start 2011/02/02 SMPK Miyata
    sql = sql & " and E020.HINBAN='" & pSpSXLSamp.hin.hinban & "' and E020.MNOREVNO=" & pSpSXLSamp.hin.mnorevno & _
                " and E020.FACTORY='" & pSpSXLSamp.hin.factory & "' and E020.OPECOND='" & pSpSXLSamp.hin.opecond & "'"
    'Add End   2011/02/02 SMPK Miyata
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getSXL = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With pSpSXLSamp
        .HSXRHWYS = rs("E018HSXRHWYS")
        .HSXRSPOT = rs("E018HSXRSPOT")
        .HSXONHWS = rs("E019HSXONHWS")
        .HSXONKWY = rs("E019HSXONKWY")
        .HSXONSPH = rs("E019HSXONSPH")
        .HSXONSPT = rs("E019HSXONSPT")
        .HSXONSPI = rs("E019HSXONSPI")
        .HSXCNHWS = rs("E019HSXCNHWS")
'        If rs("CS_FROMTO") = 1 Then
        'Cs検査頻度_位=「6」or「9」の場合、TOP/BOT保証とする。 09/01/06 ooba
        If rs("E019HSXCNKHI") = "6" Or rs("E019HSXCNKHI") = "9" Then
            .CS_FROMTO = True
        Else
            .CS_FROMTO = False
        End If
        .HSXLTSPI = rs("E019HSXLTSPI")
        .HSXLTHWS = rs("E019HSXLTHWS")
        .HSXBM1HS = rs("E020HSXBM1HS")
        .HSXBM1SH = rs("E020HSXBM1SH")
        .HSXBM1ST = rs("E020HSXBM1ST")
        .HSXBM1SR = rs("E020HSXBM1SR")
        .HSXBM1NS = rs("E020HSXBM1NS")
        .HSXBM1SZ = rs("E020HSXBM1SZ")
        .HSXBM1ET = fncNullCheck(rs("E020HSXBM1ET"))
        .HSXBM2HS = rs("E020HSXBM2HS")
        .HSXBM2SH = rs("E020HSXBM2SH")
        .HSXBM2ST = rs("E020HSXBM2ST")
        .HSXBM2SR = rs("E020HSXBM2SR")
        .HSXBM2NS = rs("E020HSXBM2NS")
        .HSXBM2SZ = rs("E020HSXBM2SZ")
        .HSXBM2ET = fncNullCheck(rs("E020HSXBM2ET"))
        .HSXBM3HS = rs("E020HSXBM3HS")
        .HSXBM3SH = rs("E020HSXBM3SH")
        .HSXBM3ST = rs("E020HSXBM3ST")
        .HSXBM3SR = rs("E020HSXBM3SR")
        .HSXBM3NS = rs("E020HSXBM3NS")
        .HSXBM3SZ = rs("E020HSXBM3SZ")
        .HSXBM3ET = fncNullCheck(rs("E020HSXBM3ET"))
        .HSXOF1HS = rs("E020HSXOF1HS")
        .HSXOF1SH = rs("E020HSXOF1SH")
        .HSXOF1ST = rs("E020HSXOF1ST")
        .HSXOF1SR = rs("E020HSXOF1SR")
        .HSXOF1NS = rs("E020HSXOF1NS")
        .HSXOF1SZ = rs("E020HSXOF1SZ")
        .HSXOF1ET = fncNullCheck(rs("E020HSXOF1ET"))
        .HSXOF2HS = rs("E020HSXOF2HS")
        .HSXOF2SH = rs("E020HSXOF2SH")
        .HSXOF2ST = rs("E020HSXOF2ST")
        .HSXOF2SR = rs("E020HSXOF2SR")
        .HSXOF2NS = rs("E020HSXOF2NS")
        .HSXOF2SZ = rs("E020HSXOF2SZ")
        .HSXOF2ET = fncNullCheck(rs("E020HSXOF2ET"))
        .HSXOF3HS = rs("E020HSXOF3HS")
        .HSXOF3SH = rs("E020HSXOF3SH")
        .HSXOF3ST = rs("E020HSXOF3ST")
        .HSXOF3SR = rs("E020HSXOF3SR")
        .HSXOF3NS = rs("E020HSXOF3NS")
        .HSXOF3SZ = rs("E020HSXOF3SZ")
        .HSXOF3ET = fncNullCheck(rs("E020HSXOF3ET"))
'        .HSXOF4HS = rs("E020HSXOF4HS")
        'OSF4保証方法_処⇒C-OSF3ﾌﾗｸﾞ　07/06/11 ooba
        If IsNull(rs("COSF3FLAG")) Then .HSXOF4HS = " " Else .HSXOF4HS = rs("COSF3FLAG")
        .HSXOF4SH = rs("E020HSXOF4SH")
        .HSXOF4ST = rs("E020HSXOF4ST")
        .HSXOF4SR = rs("E020HSXOF4SR")
        .HSXOF4NS = rs("E020HSXOF4NS")
        .HSXOF4SZ = rs("E020HSXOF4SZ")
        .HSXOF4ET = fncNullCheck(rs("E020HSXOF4ET"))
        .HSXDENHS = rs("E020HSXDENHS")
        .HSXDVDHS = rs("E020HSXDVDHS")
        .HSXLDLHS = rs("E020HSXLDLHS")
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        'Add Start 2010/12/10 SMPK Miyata
        If IsNull(rs("HSXCHS")) Then .HSXCHS = " " Else .HSXCHS = rs("HSXCHS")
        If IsNull(rs("HSXCSZ")) Then .HSXCSZ = " " Else .HSXCSZ = rs("HSXCSZ")
        If IsNull(rs("HSXCJHS")) Then .HSXCJHS = " " Else .HSXCJHS = rs("HSXCJHS")
        If IsNull(rs("HSXCJNS")) Then .HSXCJNS = "  " Else .HSXCJNS = rs("HSXCJNS")
        If IsNull(rs("HSXCJLTHS")) Then .HSXCJLTHS = " " Else .HSXCJLTHS = rs("HSXCJLTHS")
        If IsNull(rs("HSXCJLTNS")) Then .HSXCJLTNS = "  " Else .HSXCJLTNS = rs("HSXCJLTNS")
        If IsNull(rs("HSXCJ2HS")) Then .HSXCJ2HS = " " Else .HSXCJ2HS = rs("HSXCJ2HS")
        If IsNull(rs("HSXCJ2NS")) Then .HSXCJ2NS = "  " Else .HSXCJ2NS = rs("HSXCJ2NS")
        'Add End   2010/12/10 SMPK Miyata
    End With
    rs.Close

    scmzc_getSXL = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getSXL = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

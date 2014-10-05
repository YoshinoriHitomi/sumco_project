Attribute VB_Name = "SB_GetSiyou"
Option Explicit

'------------------------------------------------
' TBCME018f[^æ¾(»èp)
'------------------------------------------------

'Tv      :e[uuTBCME018v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :oR[h
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :õL[ÍA¢HINBAN£+uMNOREVNOv+uFACTORYv+uOPECONDvÌ¶ñÆ·é
'ð      :2003/09/10 VKì¬@VXeuC

Public Function funGet_TBCME018(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME018"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HSXTYPE, HSXD1CEN, HSXCDIR, HSXRMIN, HSXRMAX, HSXRAMIN, HSXRAMAX, "
    sql = sql & "HSXRMCAL, HSXRMBNP, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS "
    sql = sql & "from TBCME018 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME018 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''oÊði[·é
     With tGetRec
        .HIN.hinban = rs("HINBAN")          ' iÔ
        .HIN.mnorevno = rs("MNOREVNO")      ' »iÔüùÔ
        .HIN.factory = rs("FACTORY")        ' Hê
        .HIN.opecond = rs("OPECOND")        ' Æð
        
        .HSXTYPE = rs("HSXTYPE")                    ' irw^Cv
        .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))    ' irw¼aPS          2003/12/12 SystemBrain NullÎ
        .HSXCDIR = rs("HSXCDIR")                    ' irw»ÊûÊ
        .HSXRMIN = fncNullCheck(rs("HSXRMIN"))      ' irwäïRºÀ          2003/12/12 SystemBrain NullÎ
        .HSXRMAX = fncNullCheck(rs("HSXRMAX"))      ' irwäïRãÀ          2003/12/12 SystemBrain NullÎ
        .HSXRAMIN = fncNullCheck(rs("HSXRAMIN"))    ' irwäïR½ÏºÀ      2003/12/12 SystemBrain NullÎ
        .HSXRAMAX = fncNullCheck(rs("HSXRAMAX"))    ' irwäïR½ÏãÀ      2003/12/12 SystemBrain NullÎ
        .HSXRMCAL = rs("HSXRMCAL")                  ' irwäïRÊàvZ
        .HSXRMBNP = fncNullCheck(rs("HSXRMBNP"))    ' irwäïRÊàªz      2003/12/12 SystemBrain NullÎ
        .HSXRSPOH = rs("HSXRSPOH")                  ' irwäïRªèÊuQû
        .HSXRSPOT = rs("HSXRSPOT")                  ' irwäïRªèÊuQ_
        .HSXRSPOI = rs("HSXRSPOI")                  ' irwäïRªèÊuQÊ
        .HSXRHWYT = rs("HSXRHWYT")                  ' irwäïRÛØû@QÎ
        .HSXRHWYS = rs("HSXRHWYS")                  ' irwäïRÛØû@Q
    End With
    Set rs = Nothing

    funGet_TBCME018 = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' TBCME019f[^æ¾(»èp)
'------------------------------------------------

'Tv      :e[uuTBCME019v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :oR[h
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :õL[ÍA¢HINBAN£+uMNOREVNOv+uFACTORYv+uOPECONDvÌ¶ñÆ·é
'ð      :2003/09/10 VKì¬@VXeuC

Public Function funGet_TBCME019(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME019"

    'HSXCNKHIÇÁ 09/01/08 ooba
    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HSXONMIN, HSXONMAX, HSXONAMN, HSXONAMX, HSXONMCL, HSXONMBP, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, "
    sql = sql & "HSXCNMIN, HSXCNMAX, HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKHI, "
    sql = sql & "HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT, HSXLTHWS "
    sql = sql & "from TBCME019 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME019 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''oÊði[·é
     With tGetRec
        .HIN.hinban = rs("HINBAN")         ' iÔ
        .HIN.mnorevno = rs("MNOREVNO")     ' »iÔüùÔ
        .HIN.factory = rs("FACTORY")       ' Hê
        .HIN.opecond = rs("OPECOND")       ' Æð
        
        .HSXONMIN = fncNullCheck(rs("HSXONMIN"))        ' irw_fZxºÀ            2003/12/12 SystemBrain NullÎ
        .HSXONMAX = fncNullCheck(rs("HSXONMAX"))        ' irw_fZxãÀ            2003/12/12 SystemBrain NullÎ
        .HSXONAMN = fncNullCheck(rs("HSXONAMN"))        ' irw_fZx½ÏºÀ        2003/12/12 SystemBrain NullÎ
        .HSXONAMX = fncNullCheck(rs("HSXONAMX"))        ' irw_fZx½ÏãÀ        2003/12/12 SystemBrain NullÎ
        .HSXONMCL = rs("HSXONMCL")                      ' irw_fZxÊàvZ
        .HSXONMBP = fncNullCheck(rs("HSXONMBP"))        ' irw_fZxÊàªz        2003/12/12 SystemBrain NullÎ
        .HSXONSPH = rs("HSXONSPH")                      ' irw_fZxªèÊuQû
        .HSXONSPT = rs("HSXONSPT")                      ' irw_fZxªèÊuQ_
        .HSXONSPI = rs("HSXONSPI")                      ' irw_fZxªèÊuQÊ
        .HSXONHWT = rs("HSXONHWT")                      ' irw_fZxÛØû@QÎ
        .HSXONHWS = rs("HSXONHWS")                      ' irw_fZxÛØû@Q
        
        .HSXCNMIN = fncNullCheck(rs("HSXCNMIN"))        ' irwYfZxºÀ            2003/12/12 SystemBrain NullÎ
        .HSXCNMAX = fncNullCheck(rs("HSXCNMAX"))        ' irwYfZxãÀ            2003/12/12 SystemBrain NullÎ
        .HSXCNSPH = rs("HSXCNSPH")                      ' irwYfZxªèÊuQû
        .HSXCNSPT = rs("HSXCNSPT")                      ' irwYfZxªèÊuQ_
        .HSXCNSPI = rs("HSXCNSPI")                      ' irwYfZxªèÊuQÊ
        .HSXCNHWT = rs("HSXCNHWT")                      ' irwYfZxÛØû@QÎ
        .HSXCNHWS = rs("HSXCNHWS")                      ' irwYfZxÛØû@Q
        .HSXCNKHI = rs("HSXCNKHI")                      ' irwYfZx¸pxQÊ 09/01/08 ooba
        
        .HSXLTMIN = fncNullCheck(rs("HSXLTMIN"))        ' irwk^CºÀ            2003/12/12 SystemBrain NullÎ
        .HSXLTMAX = fncNullCheck(rs("HSXLTMAX"))        ' irwk^CãÀ            2003/12/12 SystemBrain NullÎ
        .HSXLTSPH = rs("HSXLTSPH")                      ' irwk^CªèÊuQû
        .HSXLTSPT = rs("HSXLTSPT")                      ' irwk^CªèÊuQ_
        .HSXLTSPI = rs("HSXLTSPI")                      ' irwk^CªèÊuQÊ
        .HSXLTHWT = rs("HSXLTHWT")                      ' irwk^CÛØû@QÎ
        .HSXLTHWS = rs("HSXLTHWS")                      ' irwk^CÛØû@Q
    End With
    Set rs = Nothing

    funGet_TBCME019 = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' TBCME020f[^æ¾(»èp)
'------------------------------------------------

'Tv      :e[uuTBCME020v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :oR[h
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :õL[ÍA¢HINBAN£+uMNOREVNOv+uFACTORYv+uOPECONDvÌ¶ñÆ·é
'ð      :2003/09/10 VKì¬@VXeuC

Public Function funGet_TBCME020(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME020"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1NS, "
    sql = sql & "HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS, HSXBM2NS, "
    sql = sql & "HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST, HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3NS, "
    sql = sql & "HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1NS, "
    sql = sql & "HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2NS, "
    sql = sql & "HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR, HSXOF3HT, HSXOF3HS, HSXOF3NS, "
    sql = sql & "HSXOF4AX, HSXOF4MX, HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4NS, "
    sql = sql & "HSXDENMX, HSXDENMN, HSXDENHT, HSXDENHS, HSXDENKU, "
    sql = sql & "HSXLDLMX, HSXLDLMN, HSXLDLHT, HSXLDLHS, HSXLDLKU, "
    sql = sql & "HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXDVDKU, "
    sql = sql & "HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK, HSXBMD1MBP, HSXBMD2MBP, HSXBMD3MBP "
    sql = sql & ", HSXGDPTK "   '' 2008/10/01 L/DL,OSF»èÛ¼Þ¯¸ÇÁ ADD By Systech
    'Add Start 2011/02/01 SMPK Miyata
    sql = sql & ", HSXCPK, HSXCSZ, HSXCHT, HSXCHS "
    sql = sql & ", HSXCJPK, HSXCJNS, HSXCJHT, HSXCJHS "
    sql = sql & ", HSXCJLTPK, HSXCJLTNS, HSXCJLTHT, HSXCJLTHS "
    sql = sql & ", HSXCJ2PK, HSXCJ2NS, HSXCJ2HT, HSXCJ2HS "
    'Add End   2011/02/01 SMPK Miyata
    sql = sql & "from TBCME020 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME020 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
     
    ''oÊði[·é
    With tGetRec
        .HIN.hinban = rs("HINBAN")       ' iÔ
        .HIN.mnorevno = rs("MNOREVNO")   ' »iÔüùÔ
        .HIN.factory = rs("FACTORY")     ' Hê
        .HIN.opecond = rs("OPECOND")     ' Æð
        
        .HSXBM1AN = fncNullCheck(rs("HSXBM1AN"))        ' irwalc1½ÏºÀ     2003/12/12 SystemBrain NullÎ
        .HSXBM1AX = fncNullCheck(rs("HSXBM1AX"))        ' irwalc1½ÏãÀ     2003/12/12 SystemBrain NullÎ
        .HSXBM1SH = rs("HSXBM1SH")                      ' irwalc1ªèÊuQû
        .HSXBM1ST = rs("HSXBM1ST")                      ' irwalc1ªèÊuQ_
        .HSXBM1SR = rs("HSXBM1SR")                      ' irwalc1ªèÊuQÌ
        .HSXBM1HT = rs("HSXBM1HT")                      ' irwalc1ÛØû@QÎ
        .HSXBM1HS = rs("HSXBM1HS")                      ' irwalc1ÛØû@Q
        .HSXBM1NS = rs("HSXBM1NS")                      ' irwalc1M@
        .HSXBM2AN = fncNullCheck(rs("HSXBM2AN"))        ' irwalc2½ÏºÀ     2003/12/12 SystemBrain NullÎ
        .HSXBM2AX = fncNullCheck(rs("HSXBM2AX"))        ' irwalc2½ÏãÀ     2003/12/12 SystemBrain NullÎ
        .HSXBM2SH = rs("HSXBM2SH")                      ' irwalc2ªèÊuQû
        .HSXBM2ST = rs("HSXBM2ST")                      ' irwalc2ªèÊuQ_
        .HSXBM2SR = rs("HSXBM2SR")                      ' irwalc2ªèÊuQÌ
        .HSXBM2HT = rs("HSXBM2HT")                      ' irwalc2ÛØû@QÎ
        .HSXBM2HS = rs("HSXBM2HS")                      ' irwalc2ÛØû@Q
        .HSXBM2NS = rs("HSXBM2NS")                      ' irwalc2M@
        .HSXBM3AN = fncNullCheck(rs("HSXBM3AN"))        ' irwalc3½ÏºÀ     2003/12/12 SystemBrain NullÎ
        .HSXBM3AX = fncNullCheck(rs("HSXBM3AX"))        ' irwalc3½ÏãÀ     2003/12/12 SystemBrain NullÎ
        .HSXBM3SH = rs("HSXBM3SH")                      ' irwalc3ªèÊuQû
        .HSXBM3ST = rs("HSXBM3ST")                      ' irwalc3ªèÊuQ_
        .HSXBM3SR = rs("HSXBM3SR")                      ' irwalc3ªèÊuQÌ
        .HSXBM3HT = rs("HSXBM3HT")                      ' irwalc3ÛØû@QÎ
        .HSXBM3HS = rs("HSXBM3HS")                      ' irwalc3ÛØû@Q
        .HSXBM3NS = rs("HSXBM3NS")                      ' irwalc3M@
        
        .HSXOF1AX = fncNullCheck(rs("HSXOF1AX"))        ' irwnre1½ÏãÀ     2003/12/12 SystemBrain NullÎ
        .HSXOF1MX = fncNullCheck(rs("HSXOF1MX"))        ' irwnre1ãÀ         2003/12/12 SystemBrain NullÎ
        .HSXOF1SH = rs("HSXOF1SH")                      ' irwnre1ªèÊuQû
        .HSXOF1ST = rs("HSXOF1ST")                      ' irwnre1ªèÊuQ_
        .HSXOF1SR = rs("HSXOF1SR")                      ' irwnre1ªèÊuQÌ
        .HSXOF1HT = rs("HSXOF1HT")                      ' irwnre1ÛØû@QÎ
        .HSXOF1HS = rs("HSXOF1HS")                      ' irwnre1ÛØû@Q
        .HSXOF1NS = rs("HSXOF1NS")                      ' irwnre1M@
        .HSXOF2AX = fncNullCheck(rs("HSXOF2AX"))        ' irwnre2½ÏãÀ     2003/12/12 SystemBrain NullÎ
        .HSXOF2MX = fncNullCheck(rs("HSXOF2MX"))        ' irwnre2ãÀ         2003/12/12 SystemBrain NullÎ
        .HSXOF2SH = rs("HSXOF2SH")                      ' irwnre2ªèÊuQû
        .HSXOF2ST = rs("HSXOF2ST")                      ' irwnre2ªèÊuQ_
        .HSXOF2SR = rs("HSXOF2SR")                      ' irwnre2ªèÊuQÌ
        .HSXOF2HT = rs("HSXOF2HT")                      ' irwnre2ÛØû@QÎ
        .HSXOF2HS = rs("HSXOF2HS")                      ' irwnre2ÛØû@Q
        .HSXOF2NS = rs("HSXOF2NS")                      ' irwnre2M@
        .HSXOF3AX = fncNullCheck(rs("HSXOF3AX"))        ' irwnre3½ÏãÀ     2003/12/12 SystemBrain NullÎ
        .HSXOF3MX = fncNullCheck(rs("HSXOF3MX"))        ' irwnre3ãÀ         2003/12/12 SystemBrain NullÎ
        .HSXOF3SH = rs("HSXOF3SH")                      ' irwnre3ªèÊuQû
        .HSXOF3ST = rs("HSXOF3ST")                      ' irwnre3ªèÊuQ_
        .HSXOF3SR = rs("HSXOF3SR")                      ' irwnre3ªèÊuQÌ
        .HSXOF3HT = rs("HSXOF3HT")                      ' irwnre3ÛØû@QÎ
        .HSXOF3HS = rs("HSXOF3HS")                      ' irwnre3ÛØû@Q
        .HSXOF3NS = rs("HSXOF3NS")                      ' irwnre3M@
        .HSXOF4AX = fncNullCheck(rs("HSXOF4AX"))        ' irwnre4½ÏãÀ     2003/12/12 SystemBrain NullÎ
        .HSXOF4MX = fncNullCheck(rs("HSXOF4MX"))        ' irwnre4ãÀ         2003/12/12 SystemBrain NullÎ
        .HSXOF4SH = rs("HSXOF4SH")                      ' irwnre4ªèÊuQû
        .HSXOF4ST = rs("HSXOF4ST")                      ' irwnre4ªèÊuQ_
        .HSXOF4SR = rs("HSXOF4SR")                      ' irwnre4ªèÊuQÌ
        .HSXOF4HT = rs("HSXOF4HT")                      ' irwnre4ÛØû@QÎ
        .HSXOF4HS = rs("HSXOF4HS")                      ' irwnre4ÛØû@Q
        .HSXOF4NS = rs("HSXOF4NS")                      ' irwnre4M@
        
        .HSXDENKU = rs("HSXDENKU")                      ' irwc¸L³
        .HSXDENMX = fncNullCheck(rs("HSXDENMX"))        ' irwcãÀ          2003/12/12 SystemBrain NullÎ
        .HSXDENMN = fncNullCheck(rs("HSXDENMN"))        ' irwcºÀ          2003/12/12 SystemBrain NullÎ
        .HSXDENHT = rs("HSXDENHT")                      ' irwcÛØû@QÎ
        .HSXDENHS = rs("HSXDENHS")                      ' irwcÛØû@Q
        .HSXDVDKU = rs("HSXDVDKU")                      ' irwcucQ¸L³
        .HSXDVDMX = fncNullCheck(rs("HSXDVDMXN"))       ' irwcucQãÀ        2003/12/12 SystemBrain NullÎ
        .HSXDVDMN = fncNullCheck(rs("HSXDVDMNN"))       ' irwcucQºÀ        2003/12/12 SystemBrain NullÎ
        .HSXDVDHT = rs("HSXDVDHT")                      ' irwcucQÛØû@QÎ
        .HSXDVDHS = rs("HSXDVDHS")                      ' irwcucQÛØû@Q
        .HSXLDLKU = rs("HSXLDLKU")                      ' irwk^ck¸L³
        .HSXLDLMX = fncNullCheck(rs("HSXLDLMX"))        ' irwk^ckãÀ        2003/12/12 SystemBrain NullÎ
        .HSXLDLMN = fncNullCheck(rs("HSXLDLMN"))        ' irwk^ckºÀ        2003/12/12 SystemBrain NullÎ
        .HSXLDLHT = rs("HSXLDLHT")                      ' irwk^ckÛØû@QÎ
        .HSXLDLHS = rs("HSXLDLHS")                      ' irwk^ckÛØû@Q
        
        If Not IsNull(rs("HSXOSF1PTK")) Then .HSXOSF1PTK = rs("HSXOSF1PTK")     ' irwnrePp^æª
        If Not IsNull(rs("HSXOSF2PTK")) Then .HSXOSF2PTK = rs("HSXOSF2PTK")     ' irwnreQp^æª
        If Not IsNull(rs("HSXOSF3PTK")) Then .HSXOSF3PTK = rs("HSXOSF3PTK")     ' irwnreRp^æª
        If Not IsNull(rs("HSXOSF4PTK")) Then .HSXOSF4PTK = rs("HSXOSF4PTK")     ' irwnreSp^æª
'        If Not IsNull(rs("HSXBMD1MBP")) Then .HSXBMD1MBP = rs("HSXBMD1MBP")     ' irwalcPÊàªz
'        If Not IsNull(rs("HSXBMD2MBP")) Then .HSXBMD2MBP = rs("HSXBMD2MBP")     ' irwalcQÊàªz
'        If Not IsNull(rs("HSXBMD3MBP")) Then .HSXBMD3MBP = rs("HSXBMD3MBP")     ' irwalcRÊàªz
        .HSXBMD1MBP = fncNullCheck(rs("HSXBMD1MBP"))                            ' irwalcPÊàªz    2003/12/12 SystemBrain NullÎ
        .HSXBMD2MBP = fncNullCheck(rs("HSXBMD2MBP"))                            ' irwalcQÊàªz    2003/12/12 SystemBrain NullÎ
        .HSXBMD3MBP = fncNullCheck(rs("HSXBMD3MBP"))                            ' irwalcRÊàªz    2003/12/12 SystemBrain NullÎ
        
        If Not IsNull(rs("HSXGDPTK")) Then .HSXGDPTK = rs("HSXGDPTK") Else .HSXGDPTK = " "  ' irwfcp^æª  '' 2008/10/01 L/DL,OSF»èÛ¼Þ¯¸ÇÁ ADD By Systech
    
        'Add Start 2011/02/01 SMPK Miyata
        If Not IsNull(rs("HSXCPK")) Then .HSXCPK = rs("HSXCPK")         ' irwbp^[æª
        If Not IsNull(rs("HSXCSZ")) Then .HSXCSZ = rs("HSXCSZ")         ' irwbªèð
        If Not IsNull(rs("HSXCHT")) Then .HSXCHT = rs("HSXCHT")         ' irwbÛØû@QÎ
        If Not IsNull(rs("HSXCHS")) Then .HSXCHS = rs("HSXCHS")         ' irwbÛØû@Q
        If Not IsNull(rs("HSXCJPK")) Then .HSXCJPK = rs("HSXCJPK")      ' irwbip^[æª
        If Not IsNull(rs("HSXCJNS")) Then .HSXCJNS = rs("HSXCJNS")      ' irwbiM@
        If Not IsNull(rs("HSXCJHT")) Then .HSXCJHT = rs("HSXCJHT")      ' irwbiÛØû@QÎ
        If Not IsNull(rs("HSXCJHS")) Then .HSXCJHS = rs("HSXCJHS")      ' irwbiÛØû@Q
        If Not IsNull(rs("HSXCJLTPK")) Then .HSXCJLTPK = rs("HSXCJLTPK")  ' irwbiksp^[æª
        If Not IsNull(rs("HSXCJLTNS")) Then .HSXCJLTNS = rs("HSXCJLTNS")  ' irwbiksM@
        If Not IsNull(rs("HSXCJLTHT")) Then .HSXCJLTHT = rs("HSXCJLTHT")  ' irwbiksÛØû@QÎ
        If Not IsNull(rs("HSXCJLTHS")) Then .HSXCJLTHS = rs("HSXCJLTHS")  ' irwbiksÛØû@Q
        If Not IsNull(rs("HSXCJ2PK")) Then .HSXCJ2PK = rs("HSXCJ2PK")   ' irwbiQp^[æª
        If Not IsNull(rs("HSXCJ2NS")) Then .HSXCJ2NS = rs("HSXCJ2NS")   ' irwbiQM@
        If Not IsNull(rs("HSXCJ2HT")) Then .HSXCJ2HT = rs("HSXCJ2HT")   ' irwbiQÛØû@QÎ
        If Not IsNull(rs("HSXCJ2HS")) Then .HSXCJ2HS = rs("HSXCJ2HS")   ' irwbiQÛØû@Q
        'Add End   2011/02/01 SMPK Miyata
    
    End With
    Set rs = Nothing

    funGet_TBCME020 = FUNCTION_RETURN_SUCCESS

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' TBCME036f[^æ¾(»èp)
'------------------------------------------------

'Tv      :e[uuTBCME036v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :oR[h
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :õL[ÍA¢HINBAN£+uMNOREVNOv+uFACTORYv+uOPECONDvÌ¶ñÆ·é
'ð      :2003/09/10 VKì¬@VXeuC

Public Function funGet_TBCME036(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME036"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
'C|OSF3»è@\ÇÁ 2007/04/23 M.Kaga STRAT ---
'*** UPDATE « Y.SIMIZU 2005/10/12 GD×²ÝÇÁ
'    sql = sql & "EPDUP, TOPREG, TAILREG, BTMSPRT, BLOCKHFLAG "
'    sql = sql & "EPDUP, TOPREG, TAILREG, BTMSPRT, BLOCKHFLAG,HSXGDLINE "
'*** UPDATE ª Y.SIMIZU 2005/10/12 GD×²ÝÇÁ
    sql = sql & "EPDUP, TOPREG, TAILREG, BTMSPRT, BLOCKHFLAG, HSXGDLINE, COSF3FLAG "
'C|OSF3»è@\ÇÁ 2007/04/23 M.Kaga END   ---
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sql = sql & ",NVL(HSXDKTMP,' ') HSXDKTMP "
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF»èÛ¼Þ¯¸ÇÁ ADD By Systech Start
    sql = sql & ",HSXLDLRMN, HSXLDLRMX, HWFLDLRMN, HWFLDLRMX, HSXOF1ARPTK, HSXOFARMIN, HSXOFARMAX, HSXOFARMHMX "
'' 2008/10/01 L/DL,OSF»èÛ¼Þ¯¸ÇÁ ADD By Systech End
    'Add Start 2011/02/01 SMPK Miyata
    sql = sql & ",HSXCJLTBND "
    'Add End   2011/02/01 SMPK Miyata
    sql = sql & ",LTCONVAL "        '2011/07/25 LT10¶»èÇÁÎ
    
    sql = sql & "from TBCME036 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME036 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''oÊði[·é
     With tGetRec
        .HIN.hinban = rs("HINBAN")          ' iÔ
        .HIN.mnorevno = rs("MNOREVNO")      ' »iÔüùÔ
        .HIN.factory = rs("FACTORY")        ' Hê
        .HIN.opecond = rs("OPECOND")        ' Æð
        
'        If Not IsNull(rs("EPDUP")) Then .EPDUP = rs("EPDUP")                    ' EPDãÀ
'        If Not IsNull(rs("TOPREG")) Then .TOPREG = rs("TOPREG")                 ' TOPK§
'        If Not IsNull(rs("TAILREG")) Then .TAILREG = rs("TAILREG")              ' TAILK§
'        If Not IsNull(rs("BTMSPRT")) Then .BTMSPRT = rs("BTMSPRT")              ' {gÍoK§
        .EPDUP = fncNullCheck(rs("EPDUP"))                                      ' EPDãÀ                   2003/12/12 SystemBrain NullÎ
        .TOPREG = fncNullCheck(rs("TOPREG"))                                    ' TOPK§                   2003/12/12 SystemBrain NullÎ
        .TAILREG = fncNullCheck(rs("TAILREG"))                                  ' TAILK§                  2003/12/12 SystemBrain NullÎ
        .BTMSPRT = fncNullCheck(rs("BTMSPRT"))                                  ' {gÍoK§            2003/12/12 SystemBrain NullÎ
        If Not IsNull(rs("BLOCKHFLAG")) Then .BLOCKHFLAG = rs("BLOCKHFLAG")     ' ubNPÊÛØiÔtO
    '*** UPDATE « Y.SIMIZU 2005/10/12 GD×²ÝÇÁ
        .HSXGDLINE = fncNullCheck(rs("HSXGDLINE"))
    '*** UPDATE ª Y.SIMIZU 2005/10/12 GD×²ÝÇÁ
    
'C|OSF3»è@\ÇÁ 2007/04/23 M.Kaga STRAT ---
        If IsNull(rs("COSF3FLAG")) = False Then .COSF3FLAG = rs("COSF3FLAG") Else .COSF3FLAG = " "            'C-OSF3Ì×¸Þ
'C|OSF3»è@\ÇÁ 2007/04/23 M.Kaga END   ---

'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF»èÛ¼Þ¯¸ÇÁ ADD By Systech Start
        .HSXLDLRMN = fncNullCheck(rs("HSXLDLRMN"))      ' iSXL/DLA±0ºÀ
        .HSXLDLRMX = fncNullCheck(rs("HSXLDLRMX"))      ' iSXL/DLA±0ãÀ
        .HWFLDLRMN = fncNullCheck(rs("HWFLDLRMN"))      ' iWFL/DLA±0ºÀ
        .HWFLDLRMX = fncNullCheck(rs("HWFLDLRMX"))      ' iWFL/DLA±0ãÀ
        If IsNull(rs("HSXOF1ARPTK")) = False Then .HSXOF1ARPTK = rs("HSXOF1ARPTK") Else .HSXOF1ARPTK = " "  ' iSXOSF1(ArAN)p^æª
        .HSXOFARMIN = fncNullCheck(rs("HSXOFARMIN"))    ' iSXOSF(ArAN)ºÀ
        .HSXOFARMAX = fncNullCheck(rs("HSXOFARMAX"))    ' iSXOSF(ArAN)ãÀ
        .HSXOFARMHMX = fncNullCheck(rs("HSXOFARMHMX"))  ' iSXOSF(ArAN)ÊàäãÀ
'' 2008/10/01 L/DL,OSF»èÛ¼Þ¯¸ÇÁ ADD By Systech End
        'Add Start 2011/02/01 SMPK Miyata
        .HSXCJLTBND = fncNullCheck(rs("HSXCJLTBND"))    ' iSXL/CJLToh Number(3,0)
        'Add End   2011/02/01 SMPK Miyata
        .HSXLT10MIN = fncNullCheck(rs("LTCONVAL"))      ' iSXL10¶ºÀ 2011/07/28
    
    End With
    Set rs = Nothing

    funGet_TBCME036 = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' TBCME036f[^æ¾(»èp)
'------------------------------------------------

'Tv      :e[uuTBCME036v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :oR[h
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :õL[ÍA¢HINBAN£+uMNOREVNOv+uFACTORYv+uOPECONDvÌ¶ñÆ·é
'ð      :2005/10/12 Y.SIMIZU

Public Function funGet_TBCME036_2(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME036_2"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HWFGDLINE "
    sql = sql & "from TBCME036 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME036_2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''oÊði[·é
     With tGetRec
        .HWFGDLINE = fncNullCheck(rs("HWFGDLINE"))                                      ' EPDãÀ                   2003/12/12 SystemBrain NullÎ
    End With
    Set rs = Nothing

    funGet_TBCME036_2 = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


'------------------------------------------------
' TBCME021f[^æ¾(»èp)
'------------------------------------------------

'Tv      :e[uuTBCME021v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :oR[h
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :õL[ÍA¢HINBAN£+uMNOREVNOv+uFACTORYv+uOPECONDvÌ¶ñÆ·é
'ð      :2003/09/10 VKì¬@VXeuC

Public Function funGet_TBCME021(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME021"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HWFTYPE, HWFRMIN, HWFRMAX, HWFRSPOH, HWFRSPOT, HWFRSPOI, "
    sql = sql & "HWFRHWYT, HWFRHWYS, HWFRMCAL, HWFRAMIN, HWFRAMAX, HWFRMBNP "
    sql = sql & "from TBCME021 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME021 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''oÊði[·é
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' iÔ
'        .HIN.mnorevno = rs("MNOREVNO")      ' »iÔüùÔ
'        .HIN.factory = rs("FACTORY")        ' Hê
'        .HIN.opecond = rs("OPECOND")        ' Æð
        
        .HWFTYPE = rs("HWFTYPE")                        ' ive^Cv
        .HWFRMIN = fncNullCheck(rs("HWFRMIN"))          ' iveäïRºÀ          2003/12/12 SystemBrain NullÎ
        .HWFRMAX = fncNullCheck(rs("HWFRMAX"))          ' iveäïRãÀ          2003/12/12 SystemBrain NullÎ
        .HWFRSPOH = rs("HWFRSPOH")                      ' iveäïRªèÊuQû
        .HWFRSPOT = rs("HWFRSPOT")                      ' iveäïRªèÊuQ_
        .HWFRSPOI = rs("HWFRSPOI")                      ' iveäïRªèÊuQÊ
        .HWFRHWYT = rs("HWFRHWYT")                      ' iveäïRÛØû@QÎ
        .HWFRHWYS = rs("HWFRHWYS")                      ' iveäïRÛØû@Q
        .HWFRMCAL = rs("HWFRMCAL")                      ' iveäïRÊàvZ
        .HWFRAMIN = fncNullCheck(rs("HWFRAMIN"))        ' iveäïR½ÏºÀ      2003/12/12 SystemBrain NullÎ
        .HWFRAMAX = fncNullCheck(rs("HWFRAMAX"))        ' iveäïR½ÏãÀ      2003/12/12 SystemBrain NullÎ
        .HWFRMBNP = fncNullCheck(rs("HWFRMBNP"))        ' iveäïRÊàªz      2003/12/12 SystemBrain NullÎ
    End With
    Set rs = Nothing

    funGet_TBCME021 = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' TBCME024f[^æ¾(»èp)
'------------------------------------------------

'Tv      :e[uuTBCME024v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :oR[h
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :õL[ÍA¢HINBAN£+uMNOREVNOv+uFACTORYv+uOPECONDvÌ¶ñÆ·é
'ð      :2003/09/10 VKì¬@VXeuC

Public Function funGet_TBCME024(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME024"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HWFMKMIN, HWFMKMAX, HWFMKSPH, HWFMKSPT, HWFMKSPR, HWFMKHWT, HWFMKHWS "
    sql = sql & "from TBCME024 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME024 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''oÊði[·é
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' iÔ
'        .HIN.mnorevno = rs("MNOREVNO")      ' »iÔüùÔ
'        .HIN.factory = rs("FACTORY")        ' Hê
'        .HIN.opecond = rs("OPECOND")        ' Æð
        
        .HWFMKMIN = fncNullCheck(rs("HWFMKMIN"))        ' ive³×wºÀ            2003/12/12 SystemBrain NullÎ
        .HWFMKMAX = fncNullCheck(rs("HWFMKMAX"))        ' ive³×wãÀ            2003/12/12 SystemBrain NullÎ
        .HWFMKSPH = rs("HWFMKSPH")                      ' ive³×wªèÊuQû
        .HWFMKSPT = rs("HWFMKSPT")                      ' ive³×wªèÊuQ_
        .HWFMKSPR = rs("HWFMKSPR")                      ' ive³×wªèÊuQÌ
        .HWFMKHWT = rs("HWFMKHWT")                      ' ive³×wÛØû@QÎ
        .HWFMKHWS = rs("HWFMKHWS")                      ' ive³×wÛØû@Q
    End With
    Set rs = Nothing

    funGet_TBCME024 = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' TBCME025f[^æ¾
'------------------------------------------------

'Tv      :e[uuTBCME025v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :oR[h
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :õL[ÍA¢HINBAN£+uMNOREVNOv+uFACTORYv+uOPECONDvÌ¶ñÆ·é
'ð      :2003/09/10 VKì¬@VXeuC

Public Function funGet_TBCME025(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME025"

    sql = "select E025.HINBAN, E025.MNOREVNO, E025.FACTORY, E025.OPECOND, "
    sql = sql & "E025.HWFONMIN, E025.HWFONMAX, E025.HWFONSPH, E025.HWFONSPT, E025.HWFONSPI, E025.HWFONHWT, E025.HWFONHWS, "
    sql = sql & "HSXONSPT, HSXONSPI, "
    sql = sql & "E025.HWFONMCL, E025.HWFONMBP, E025.HWFONAMN, E025.HWFONAMX, "
    sql = sql & "E025.HWFOS1MN, E025.HWFOS1MX, E025.HWFOS1SH, E025.HWFOS1ST, E025.HWFOS1SI, E025.HWFOS1HT, E025.HWFOS1HS, E025.HWFOS1NS, "
    sql = sql & "E025.HWFOS2MN, E025.HWFOS2MX, E025.HWFOS2SH, E025.HWFOS2ST, E025.HWFOS2SI, E025.HWFOS2HT, E025.HWFOS2HS, E025.HWFOS2NS, "
    sql = sql & "E025.HWFOS3MN, E025.HWFOS3MX, E025.HWFOS3SH, E025.HWFOS3ST, E025.HWFOS3SI, E025.HWFOS3HT, E025.HWFOS3HS, E025.HWFOS3NS, "
    ''c¶_fdlæ¾ÇÁ@03/12/09 ooba
    sql = sql & "E025.HWFZOMIN, E025.HWFZOMAX, E025.HWFZOSPH, E025.HWFZOSPT, E025.HWFZOSPI, E025.HWFZOHWT, E025.HWFZOHWS, E025.HWFZONSW, "
    sql = sql & "E025.HWFANTNP, E025.HWFANTIM "
    sql = sql & "from TBCME025 E025, TBCME019 E019 "
    sql = sql & "Where E025.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E025.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E025.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E025.OPECOND = '" & tHIN.opecond & "' and "
    sql = sql & "      E019.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E019.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E019.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E019.OPECOND = '" & tHIN.opecond & "'"
    
    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME025 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''oÊði[·é
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' iÔ
'        .HIN.mnorevno = rs("MNOREVNO")      ' »iÔüùÔ
'        .HIN.factory = rs("FACTORY")        ' Hê
'        .HIN.opecond = rs("OPECOND")        ' Æð
        
        .HWFONMIN = fncNullCheck(rs("HWFONMIN"))        ' ive_fZxºÀ            2003/12/12 SystemBrain NullÎ
        .HWFONMAX = fncNullCheck(rs("HWFONMAX"))        ' ive_fZxãÀ            2003/12/12 SystemBrain NullÎ
        .HWFONSPH = rs("HWFONSPH")                      ' ive_fZxªèÊuQû
'        .HWFONSPT = rs("HWFONSPT")                      ' ive_fZxªèÊuQ_
'        .HWFONSPI = rs("HWFONSPI")                      ' ive_fZxªèÊuQÊ
        .HWFONSPT = rs("HSXONSPT")                      ' irw_fZxªèÊuQ_
        .HWFONSPI = rs("HSXONSPI")                      ' irw_fZxªèÊuQÊ
        .HWFONHWT = rs("HWFONHWT")                      ' ive_fZxÛØû@QÎ
        .HWFONHWS = rs("HWFONHWS")                      ' ive_fZxÛØû@Q
        .HWFONMCL = rs("HWFONMCL")                      ' ive_fZxÊàvZ
        .HWFONMBP = fncNullCheck(rs("HWFONMBP"))        ' ive_fZxÊàªz        2003/12/12 SystemBrain NullÎ
        .HWFONAMN = fncNullCheck(rs("HWFONAMN"))        ' ive_fZx½ÏºÀ        2003/12/12 SystemBrain NullÎ
        .HWFONAMX = fncNullCheck(rs("HWFONAMX"))        ' ive_fZx½ÏãÀ        2003/12/12 SystemBrain NullÎ
        
        .HWFOS1MN = fncNullCheck(rs("HWFOS1MN"))        ' ive_fÍoPºÀ          2003/12/12 SystemBrain NullÎ
        .HWFOS1MX = fncNullCheck(rs("HWFOS1MX"))        ' ive_fÍoPãÀ          2003/12/12 SystemBrain NullÎ
        .HWFOS1SH = rs("HWFOS1SH")                      ' ive_fÍoPªèÊuQû
        .HWFOS1ST = rs("HWFOS1ST")                      ' ive_fÍoPªèÊuQ_
        .HWFOS1SI = rs("HWFOS1SI")                      ' ive_fÍoPªèÊuQÊ
        .HWFOS1HT = rs("HWFOS1HT")                      ' ive_fÍoPÛØû@QÎ
        .HWFOS1HS = rs("HWFOS1HS")                      ' ive_fÍoPÛØû@Q
        .HWFOS1NS = rs("HWFOS1NS")                      ' ive_fÍoPM@
        
        .HWFOS2MN = fncNullCheck(rs("HWFOS2MN"))        ' ive_fÍoQºÀ          2003/12/12 SystemBrain NullÎ
        .HWFOS2MX = fncNullCheck(rs("HWFOS2MX"))        ' ive_fÍoQãÀ          2003/12/12 SystemBrain NullÎ
        .HWFOS2SH = rs("HWFOS2SH")                      ' ive_fÍoQªèÊuQû
        .HWFOS2ST = rs("HWFOS2ST")                      ' ive_fÍoQªèÊuQ_
        .HWFOS2SI = rs("HWFOS2SI")                      ' ive_fÍoQªèÊuQÊ
        .HWFOS2HT = rs("HWFOS2HT")                      ' ive_fÍoQÛØû@QÎ
        .HWFOS2HS = rs("HWFOS2HS")                      ' ive_fÍoQÛØû@Q
        .HWFOS2NS = rs("HWFOS2NS")                      ' ive_fÍoQM@
        
        .HWFOS3MN = fncNullCheck(rs("HWFOS3MN"))        ' ive_fÍoRºÀ          2003/12/12 SystemBrain NullÎ
        .HWFOS3MX = fncNullCheck(rs("HWFOS3MX"))        ' ive_fÍoRãÀ          2003/12/12 SystemBrain NullÎ
        .HWFOS3SH = rs("HWFOS3SH")                      ' ive_fÍoRªèÊuQû
        .HWFOS3ST = rs("HWFOS3ST")                      ' ive_fÍoRªèÊuQ_
        .HWFOS3SI = rs("HWFOS3SI")                      ' ive_fÍoRªèÊuQÊ
        .HWFOS3HT = rs("HWFOS3HT")                      ' ive_fÍoRÛØû@QÎ
        .HWFOS3HS = rs("HWFOS3HS")                      ' ive_fÍoRÛØû@Q
        .HWFOS3NS = rs("HWFOS3NS")                      ' ive_fÍoRM@
        
        ''c¶_fdlæ¾ÇÁ@03/12/09 ooba START ==============================>
'''        If IsNull(rs("HWFZOMIN")) = False Then .HWFZOMIN = rs("HWFZOMIN") ' ivec¶_fºÀ
'''        If IsNull(rs("HWFZOMAX")) = False Then .HWFZOMAX = rs("HWFZOMAX") ' ivec¶_fãÀ
'''        .HWFZOSPH = rs("HWFZOSPH")                  ' ivec¶_fªèÊuQû
'''        .HWFZOSPT = rs("HWFZOSPT")                  ' ivec¶_fªèÊuQ_
'''        .HWFZOSPI = rs("HWFZOSPI")                  ' ivec¶_fªèÊuQÊ
'''        .HWFZOHWT = rs("HWFZOHWT")                  ' ivec¶_fÛØû@QÎ
'''        .HWFZOHWS = rs("HWFZOHWS")                  ' ivec¶_fÛØû@Q
'''        .HWFZONSW = rs("HWFZONSW")                  ' ivec¶_fM@

        .HWFZOMIN = fncNullCheck(rs("HWFZOMIN"))    ' ivec¶_fºÀ
        .HWFZOMAX = fncNullCheck(rs("HWFZOMAX"))    ' ivec¶_fãÀ
        If IsNull(rs("HWFZOSPH")) = False Then .HWFZOSPH = rs("HWFZOSPH") ' ivec¶_fªèÊuQû
        If IsNull(rs("HWFZOSPT")) = False Then .HWFZOSPT = rs("HWFZOSPT") ' ivec¶_fªèÊuQ_
        If IsNull(rs("HWFZOSPI")) = False Then .HWFZOSPI = rs("HWFZOSPI") ' ivec¶_fªèÊuQÊ
        If IsNull(rs("HWFZOHWT")) = False Then .HWFZOHWT = rs("HWFZOHWT") ' ivec¶_fÛØû@QÎ
        If IsNull(rs("HWFZOHWS")) = False Then .HWFZOHWS = rs("HWFZOHWS") ' ivec¶_fÛØû@Q
        If IsNull(rs("HWFZONSW")) = False Then .HWFZONSW = rs("HWFZONSW") ' ivec¶_fM@
        ''c¶_fdlæ¾ÇÁ@03/12/09 ooba END ================================>
        
        .HWFANTIM = fncNullCheck(rs("HWFANTIM"))        ' ive`mÔ                2003/12/12 SystemBrain NullÎ
        .HWFANTNP = fncNullCheck(rs("HWFANTNP"))        ' ive`m·x                2003/12/12 SystemBrain NullÎ
    End With
    Set rs = Nothing

    funGet_TBCME025 = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' TBCME026f[^æ¾
'------------------------------------------------

'Tv      :e[uuTBCME026v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :oR[h
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :õL[ÍA¢HINBAN£+uMNOREVNOv+uFACTORYv+uOPECONDvÌ¶ñÆ·é
'ð      :2003/09/10 VKì¬@VXeuC

Public Function funGet_TBCME026(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME026"

    'DSODÊßÀ°Ýæªæ¾ÇÁ@04/08/09
    'GDdlæ¾ÇÁ@05/01/26
''    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
''    sql = sql & "HWFDSOPTK, "
''    sql = sql & "HWFDENKU, HWFDENMX, HWFDENMN, HWFDENHT, HWFDENHS, "
''    sql = sql & "HWFDVDKU, HWFDVDMXN, HWFDVDMNN, HWFDVDHT, HWFDVDHS, "
''    sql = sql & "HWFLDLKU, HWFLDLMX, HWFLDLMN, HWFLDLHT, HWFLDLHS, "
''    sql = sql & "HWFGDSPH, HWFGDSPT, HWFGDSPR, "
''    sql = sql & "HWFDSOMX, HWFDSOMN, HWFDSOAX, HWFDSOAN, HWFDSOHT, HWFDSOHS "
''    sql = sql & "from TBCME026 "
''    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
''    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
''    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
''    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    'DK±Æ°Ù·xÇÁ@06/12/22 ooba
    sql = "select E025.HINBAN, E025.MNOREVNO, E025.FACTORY, E025.OPECOND, "
    sql = sql & "E026.HWFDENKU, E026.HWFDENMX, E026.HWFDENMN, E026.HWFDENHT, E026.HWFDENHS, "
    sql = sql & "E026.HWFDVDKU, E026.HWFDVDMXN, E026.HWFDVDMNN, E026.HWFDVDHT, E026.HWFDVDHS, "
    sql = sql & "E026.HWFLDLKU, E026.HWFLDLMX, E026.HWFLDLMN, E026.HWFLDLHT, E026.HWFLDLHS, "
    sql = sql & "E026.HWFGDSPH, E026.HWFGDSPT, E026.HWFGDSPR, "
    sql = sql & "E026.HWFDSOMX, E026.HWFDSOMN, E026.HWFDSOAX, E026.HWFDSOAN, E026.HWFDSOHT, "
    sql = sql & "E026.HWFDSOHS, E026.HWFDSOPTK, E025.HWFANTNP "
    sql = sql & ",E026.HWFGDPTK "    '' 2008/10/01 L/DL,OSF»èÛ¼Þ¯¸ÇÁ ADD By Systech
    sql = sql & "from TBCME025 E025, TBCME026 E026 "
    sql = sql & "Where E025.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E025.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E025.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E025.OPECOND = '" & tHIN.opecond & "' and "
    sql = sql & "      E026.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E026.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E026.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E026.OPECOND = '" & tHIN.opecond & "'"
    
    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME026 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''oÊði[·é
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' iÔ
'        .HIN.mnorevno = rs("MNOREVNO")      ' »iÔüùÔ
'        .HIN.factory = rs("FACTORY")        ' Hê
'        .HIN.opecond = rs("OPECOND")        ' Æð
        
        .HWFDSOMX = fncNullCheck(rs("HWFDSOMX"))        ' ivecrncãÀ            2003/12/12 SystemBrain NullÎ
        .HWFDSOMN = fncNullCheck(rs("HWFDSOMN"))        ' ivecrncºÀ            2003/12/12 SystemBrain NullÎ
        .HWFDSOAX = fncNullCheck(rs("HWFDSOAX"))        ' ivecrncÌæãÀ        2003/12/12 SystemBrain NullÎ
        .HWFDSOAN = fncNullCheck(rs("HWFDSOAN"))        ' ivecrncÌæºÀ        2003/12/12 SystemBrain NullÎ
        .HWFDSOHT = rs("HWFDSOHT")                      ' ivecrncÛØû@QÎ
        .HWFDSOHS = rs("HWFDSOHS")                      ' ivecrncÛØû@Q
        If IsNull(rs("HWFDSOPTK")) = False Then .HWFDSOPTK = rs("HWFDSOPTK") Else .HWFDSOPTK = " "          'p^[æª@04/08/09 ooba
        
        ''GDdlæ¾ÇÁ@05/01/26 ooba START ========================================>
        .HWFDENKU = rs("HWFDENKU")                      ' ivec¸L³
        .HWFDENMX = fncNullCheck(rs("HWFDENMX"))        ' ivecãÀ
        .HWFDENMN = fncNullCheck(rs("HWFDENMN"))        ' ivecºÀ
        .HWFDENHT = rs("HWFDENHT")                      ' ivecÛØû@QÎ
        .HWFDENHS = rs("HWFDENHS")                      ' ivecÛØû@Q
        .HWFDVDKU = rs("HWFDVDKU")                      ' ivecucQ¸L³
        .HWFDVDMXN = fncNullCheck(rs("HWFDVDMXN"))      ' ivecucQãÀ
        .HWFDVDMNN = fncNullCheck(rs("HWFDVDMNN"))      ' ivecucQºÀ
        .HWFDVDHT = rs("HWFDVDHT")                      ' ivecucQÛØû@QÎ
        .HWFDVDHS = rs("HWFDVDHS")                      ' ivecucQÛØû@Q
        .HWFLDLKU = rs("HWFLDLKU")                      ' ivek^ck¸L³
        .HWFLDLMX = fncNullCheck(rs("HWFLDLMX"))        ' ivek^ckãÀ
        .HWFLDLMN = fncNullCheck(rs("HWFLDLMN"))        ' ivek^ckºÀ
        .HWFLDLHT = rs("HWFLDLHT")                      ' ivek^ckÛØû@QÎ
        .HWFLDLHS = rs("HWFLDLHS")                      ' ivek^ckÛØû@Q
        .HWFGDSPH = rs("HWFGDSPH")                      ' ivefcªèÊuQû
        .HWFGDSPT = rs("HWFGDSPT")                      ' ivefcªèÊuQ_
        .HWFGDSPR = rs("HWFGDSPR")                      ' ivefcªèÊuQÌ
        ''GDdlæ¾ÇÁ@05/01/26 ooba END ==========================================>
        
        If Not IsNull(rs("HWFANTNP")) Then .HWFANTNP = rs("HWFANTNP")   ' ive`m·x@06/12/22 ooba
        
        If Not IsNull(rs("HWFGDPTK")) Then .HWFGDPTK = rs("HWFGDPTK") Else .HWFGDPTK = " "  ' ivefcp^æª  '' 2008/10/01 L/DL,OSF»èÛ¼Þ¯¸ÇÁ ADD By Systech
    End With
    Set rs = Nothing

    funGet_TBCME026 = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' TBCME028f[^æ¾
'------------------------------------------------

'Tv      :e[uuTBCME028v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :oR[h
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :õL[ÍA¢HINBAN£+uMNOREVNOv+uFACTORYv+uOPECONDvÌ¶ñÆ·é
'ð      :2003/09/10 VKì¬@VXeuC

Public Function funGet_TBCME028(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME028"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "

''Upd start 2005/06/28 (TCS)T.terauchi  SPV9_Î
''    sql = sql & "HWFSPVMX, HWFSPVSH, HWFSPVST, HWFSPVSI, HWFSPVHT, HWFSPVHS, "
    sql = sql & "HWFSPVMX, HWFSPVMXN, HWFSPVSH, HWFSPVST, HWFSPVSI, HWFSPVHT, HWFSPVHS, "
    sql = sql & "HWFSPVKN, HWFDLKHN, "
''Upd end   2005/06/28 (TCS)T.Terauchi  SPV9_Î

'«ÇÁ SPV»èÇÁ 2006/06/12 SMP)kondoh ---------------
    sql = sql & "HWFSPVAMN, "
'ªÇÁ SPV»èÇÁ 2006/06/12 SMP)kondoh ---------------
    
    sql = sql & "HWFDLMIN, HWFDLMAX, HWFDLSPH, HWFDLSPT, HWFDLSPI, HWFDLHWT, HWFDLHWS "
    sql = sql & "from TBCME028 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME028 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''oÊði[·é
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' iÔ
'        .HIN.mnorevno = rs("MNOREVNO")      ' »iÔüùÔ
'        .HIN.factory = rs("FACTORY")        ' Hê
'        .HIN.opecond = rs("OPECOND")        ' Æð
        
    ''Upd start 2005/06/28 (TCS)T.Terauchi  SPV9_Î
    ''    .HWFSPVMX = fncNullCheck(rs("HWFSPVMX"))        ' iverouedãÀ          2003/12/12 SystemBrain NullÎ
        .HWFSPVMX = fncNullCheck(rs("HWFSPVMXN"))       ' iverouedãÀ
        .HWFSPVKN = rs("HWFSPVKN")                      ' iveroued¸pxQ²
        .HWFDLKHN = rs("HWFDLKHN")                      ' ivegU·¸pxQ²
    ''Upd end   2005/06/28 (TCS)T.Terauchi  SPV9_Î
    
'«ÇÁ SPV»èÇÁ 2006/06/12 SMP)kondoh ---------------
        .HWFSPVAM = fncNullCheck(rs("HWFSPVAMN"))       ' iveroued½Ï
'ªÇÁ SPV»èÇÁ 2006/06/12 SMP)kondoh ---------------
        
        
        .HWFSPVSH = rs("HWFSPVSH")                      ' iverouedªèÊuQû
        .HWFSPVST = rs("HWFSPVST")                      ' iverouedªèÊuQ_
        .HWFSPVSI = rs("HWFSPVSI")                      ' iverouedªèÊuQÊ
        .HWFSPVHT = rs("HWFSPVHT")                      ' iverouedÛØû@QÎ
        .HWFSPVHS = rs("HWFSPVHS")                      ' iverouedÛØû@Q
        
        .HWFDLMIN = fncNullCheck(rs("HWFDLMIN"))        ' ivegU·ºÀ              2003/12/12 SystemBrain NullÎ
        .HWFDLMAX = fncNullCheck(rs("HWFDLMAX"))        ' ivegU·ãÀ              2003/12/12 SystemBrain NullÎ
        .HWFDLSPH = rs("HWFDLSPH")                      ' ivegU·ªèÊuQû
        .HWFDLSPT = rs("HWFDLSPT")                      ' ivegU·ªèÊuQ_
        .HWFDLSPI = rs("HWFDLSPI")                      ' ivegU·ªèÊuQÊ
        .HWFDLHWT = rs("HWFDLHWT")                      ' ivegU·ÛØû@QÎ
        .HWFDLHWS = rs("HWFDLHWS")                      ' ivegU·ÛØû@Q
    End With
    Set rs = Nothing

    funGet_TBCME028 = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' TBCME029f[^æ¾
'------------------------------------------------

'Tv      :e[uuTBCME029v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :oR[h
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :õL[ÍA¢HINBAN£+uMNOREVNOv+uFACTORYv+uOPECONDvÌ¶ñÆ·é
'ð      :2003/09/10 VKì¬@VXeuC

Public Function funGet_TBCME029(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME029"

'«ÏX M»fÇÁ 2006/02/15 SMPÎì ---------------
    'AN·x`FbNÌ×ÉTBCME025©çive`m·xðæ¾·é
'    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
'    sql = sql & "HWFOF1AX, HWFOF1MX, HWFOF1SH, HWFOF1ST, HWFOF1SR, HWFOF1HT, HWFOF1HS, HWFOF1NS, HWFOF1ET, HWFOF1SZ, "
'    sql = sql & "HWFOF2AX, HWFOF2MX, HWFOF2SH, HWFOF2ST, HWFOF2SR, HWFOF2HT, HWFOF2HS, HWFOF2NS, HWFOF2ET, HWFOF2SZ, "
'    sql = sql & "HWFOF3AX, HWFOF3MX, HWFOF3SH, HWFOF3ST, HWFOF3SR, HWFOF3HT, HWFOF3HS, HWFOF3NS, HWFOF3ET, HWFOF3SZ, "
'    sql = sql & "HWFOF4AX, HWFOF4MX, HWFOF4SH, HWFOF4ST, HWFOF4SR, HWFOF4HT, HWFOF4HS, HWFOF4NS, HWFOF4ET, HWFOF4SZ, "
'    sql = sql & "HWFBM1AN, HWFBM1AX, HWFBM1SH, HWFBM1ST, HWFBM1SR, HWFBM1HT, HWFBM1HS, HWFBM1NS, HWFBM1ET, HWFBM1SZ, "
'    sql = sql & "HWFBM2AN, HWFBM2AX, HWFBM2SH, HWFBM2ST, HWFBM2SR, HWFBM2HT, HWFBM2HS, HWFBM2NS, HWFBM2ET, HWFBM2SZ, "
'    sql = sql & "HWFBM3AN, HWFBM3AX, HWFBM3SH, HWFBM3ST, HWFBM3SR, HWFBM3HT, HWFBM3HS, HWFBM3NS, HWFBM3ET, HWFBM3SZ, "
'    sql = sql & "HWFOSF1PTK, HWFOSF2PTK, HWFOSF3PTK, HWFOSF4PTK, "
'    sql = sql & "HWFBM1MBP, HWFBM2MBP, HWFBM3MBP, HWFBM1MCL, HWFBM2MCL, HWFBM3MCL "
'    sql = sql & "from TBCME029 "
'    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
'    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
'    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
'    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    sql = "select E029.HINBAN, E029.MNOREVNO, E029.FACTORY, E029.OPECOND, "
    sql = sql & "E029.HWFOF1AX, E029.HWFOF1MX, E029.HWFOF1SH, E029.HWFOF1ST, E029.HWFOF1SR, E029.HWFOF1HT, E029.HWFOF1HS, E029.HWFOF1NS, E029.HWFOF1ET, HWFOF1SZ, "
    sql = sql & "E029.HWFOF2AX, E029.HWFOF2MX, E029.HWFOF2SH, E029.HWFOF2ST, E029.HWFOF2SR, E029.HWFOF2HT, E029.HWFOF2HS, E029.HWFOF2NS, E029.HWFOF2ET, HWFOF2SZ, "
    sql = sql & "E029.HWFOF3AX, E029.HWFOF3MX, E029.HWFOF3SH, E029.HWFOF3ST, E029.HWFOF3SR, E029.HWFOF3HT, E029.HWFOF3HS, E029.HWFOF3NS, E029.HWFOF3ET, HWFOF3SZ, "
    sql = sql & "E029.HWFOF4AX, E029.HWFOF4MX, E029.HWFOF4SH, E029.HWFOF4ST, E029.HWFOF4SR, E029.HWFOF4HT, E029.HWFOF4HS, E029.HWFOF4NS, E029.HWFOF4ET, HWFOF4SZ, "
    sql = sql & "E029.HWFBM1AN, E029.HWFBM1AX, E029.HWFBM1SH, E029.HWFBM1ST, E029.HWFBM1SR, E029.HWFBM1HT, E029.HWFBM1HS, E029.HWFBM1NS, E029.HWFBM1ET, HWFBM1SZ, "
    sql = sql & "E029.HWFBM2AN, E029.HWFBM2AX, E029.HWFBM2SH, E029.HWFBM2ST, E029.HWFBM2SR, E029.HWFBM2HT, E029.HWFBM2HS, E029.HWFBM2NS, E029.HWFBM2ET, HWFBM2SZ, "
    sql = sql & "E029.HWFBM3AN, E029.HWFBM3AX, E029.HWFBM3SH, E029.HWFBM3ST, E029.HWFBM3SR, E029.HWFBM3HT, E029.HWFBM3HS, E029.HWFBM3NS, E029.HWFBM3ET, HWFBM3SZ, "
    sql = sql & "E029.HWFOSF1PTK, E029.HWFOSF2PTK, E029.HWFOSF3PTK, E029.HWFOSF4PTK, "
    sql = sql & "E029.HWFBM1MBP, E029.HWFBM2MBP, E029.HWFBM3MBP, E029.HWFBM1MCL, E029.HWFBM2MCL, E029.HWFBM3MCL, "
    sql = sql & "E025.HWFANTNP "
'--- 2010/01/20 SIRDÎ SPK habuki ADD START(OSF4->SIRD)
    sql = sql & ",E048.HWFSIRDMX "          '²ó]ÊãÀ
    sql = sql & ",E048.HWFSIRDHT "          '²ó]ÊÛØû@QÎ
    sql = sql & ",E048.HWFSIRDHS "          '²ó]ÊÛØû@Q
    sql = sql & ",E048.HWFSIRDSZ "          '²ó]Êªèð
'--- 2010/01/20 SIRDÎ SPK habuki ADD  END (OSF4->SIRD)
    sql = sql & "from TBCME029 E029 "
    sql = sql & "    ,TBCME025 E025 "
'--- 2010/01/20 SIRDÎ SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "    ,TBCME048 E048 "
'--- 2010/01/20 SIRDÎ SPK habuki ADD  END (OSF4->SIRD)
    sql = sql & "Where E029.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E029.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E029.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E029.OPECOND = '" & tHIN.opecond & "' and "
    sql = sql & "      E025.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E025.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E025.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E025.OPECOND = '" & tHIN.opecond & "'"
'ªÏX M»fÇÁ 2006/02/15 SMPÎì ---------------
'--- 2010/01/20 SIRDÎ SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "  and E048.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E048.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E048.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E048.OPECOND = '" & tHIN.opecond & "'"
'--- 2010/01/20 SIRDÎ SPK habuki ADD  END (OSF4->SIRD)
    
    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME029 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''oÊði[·é
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' iÔ
'        .HIN.mnorevno = rs("MNOREVNO")      ' »iÔüùÔ
'        .HIN.factory = rs("FACTORY")        ' Hê
'        .HIN.opecond = rs("OPECOND")        ' Æð
        
        .HWFOF1AX = fncNullCheck(rs("HWFOF1AX"))        ' ivenreP½ÏãÀ        2003/12/12 SystemBrain NullÎ
        .HWFOF1MX = fncNullCheck(rs("HWFOF1MX"))        ' ivenrePãÀ            2003/12/12 SystemBrain NullÎ
        .HWFOF1SH = rs("HWFOF1SH")                      ' ivenrePªèÊuQû
        .HWFOF1ST = rs("HWFOF1ST")                      ' ivenrePªèÊuQ_
        .HWFOF1SR = rs("HWFOF1SR")                      ' ivenrePªèÊuQÌ
        .HWFOF1HT = rs("HWFOF1HT")                      ' ivenrePÛØû@QÎ
        .HWFOF1HS = rs("HWFOF1HS")                      ' ivenrePÛØû@Q
        .HWFOF1NS = rs("HWFOF1NS")                      ' ivenrePM@
        .HWFOF1ET = fncNullCheck(rs("HWFOF1ET"))        ' ivenrePIðdsã      2003/12/12 SystemBrain NullÎ
        .HWFOF1SZ = rs("HWFOF1SZ")                      ' ivenrePªèð
        .HWFOF2AX = fncNullCheck(rs("HWFOF2AX"))        ' ivenreQ½ÏãÀ        2003/12/12 SystemBrain NullÎ
        .HWFOF2MX = fncNullCheck(rs("HWFOF2MX"))        ' ivenreQãÀ            2003/12/12 SystemBrain NullÎ
        .HWFOF2SH = rs("HWFOF2SH")                      ' ivenreQªèÊuQû
        .HWFOF2ST = rs("HWFOF2ST")                      ' ivenreQªèÊuQ_
        .HWFOF2SR = rs("HWFOF2SR")                      ' ivenreQªèÊuQÌ
        .HWFOF2HT = rs("HWFOF2HT")                      ' ivenreQÛØû@QÎ
        .HWFOF2HS = rs("HWFOF2HS")                      ' ivenreQÛØû@Q
        .HWFOF2NS = rs("HWFOF2NS")                      ' ivenreQM@
        .HWFOF2ET = fncNullCheck(rs("HWFOF2ET"))        ' ivenreQIðdsã      2003/12/12 SystemBrain NullÎ
        .HWFOF2SZ = rs("HWFOF2SZ")                      ' ivenreQªèð
        .HWFOF3AX = fncNullCheck(rs("HWFOF3AX"))        ' ivenreR½ÏãÀ        2003/12/12 SystemBrain NullÎ
        .HWFOF3MX = fncNullCheck(rs("HWFOF3MX"))        ' ivenreRãÀ            2003/12/12 SystemBrain NullÎ
        .HWFOF3SH = rs("HWFOF3SH")                      ' ivenreRªèÊuQû
        .HWFOF3ST = rs("HWFOF3ST")                      ' ivenreRªèÊuQ_
        .HWFOF3SR = rs("HWFOF3SR")                      ' ivenreRªèÊuQÌ
        .HWFOF3HT = rs("HWFOF3HT")                      ' ivenreRÛØû@QÎ
        .HWFOF3HS = rs("HWFOF3HS")                      ' ivenreRÛØû@Q
        .HWFOF3NS = rs("HWFOF3NS")                      ' ivenreRM@
        .HWFOF3ET = fncNullCheck(rs("HWFOF3ET"))        ' ivenreRIðdsã      2003/12/12 SystemBrain NullÎ
        .HWFOF3SZ = rs("HWFOF3SZ")                      ' ivenreRªèð
        .HWFOF4AX = fncNullCheck(rs("HWFOF4AX"))        ' ivenreS½ÏãÀ        2003/12/12 SystemBrain NullÎ
        .HWFOF4MX = fncNullCheck(rs("HWFOF4MX"))        ' ivenreSãÀ            2003/12/12 SystemBrain NullÎ
        .HWFOF4SH = rs("HWFOF4SH")                      ' ivenreSªèÊuQû
        .HWFOF4ST = rs("HWFOF4ST")                      ' ivenreSªèÊuQ_
        .HWFOF4SR = rs("HWFOF4SR")                      ' ivenreSªèÊuQÌ
        .HWFOF4HT = rs("HWFOF4HT")                      ' ivenreSÛØû@QÎ
        .HWFOF4HS = rs("HWFOF4HS")                      ' ivenreSÛØû@Q
        .HWFOF4NS = rs("HWFOF4NS")                      ' ivenreSM@
        .HWFOF4ET = fncNullCheck(rs("HWFOF4ET"))        ' ivenreSIðdsã      2003/12/12 SystemBrain NullÎ
        .HWFOF4SZ = rs("HWFOF4SZ")                      ' ivenreSªèð
'--- 2010/01/20 SIRDÎ SPK habuki ADD START(OSF4->SIRD)
        If IsNull(rs("HWFSIRDMX")) = False Then .HWFOF4MX = rs("HWFSIRDMX") Else .HWFOF4MX = "0"        ' ²ó]ÊãÀ
        If IsNull(rs("HWFSIRDHT")) = False Then .HWFOF4HT = rs("HWFSIRDHT") Else .HWFOF4HT = " "        ' ²ó]ÊÛØû@QÎ
        If IsNull(rs("HWFSIRDHS")) = False Then .HWFOF4HS = rs("HWFSIRDHS") Else .HWFOF4HS = " "        ' ²ó]ÊÛØû@Q
        If IsNull(rs("HWFSIRDSZ")) = False Then .HWFOF4SZ = rs("HWFSIRDSZ") Else .HWFOF4SZ = " "        ' ²ó]Êªèð
'--- 2010/01/20 SIRDÎ SPK habuki ADD  END (OSF4->SIRD)
        
        .HWFBM1AN = fncNullCheck(rs("HWFBM1AN"))        ' ivealcP½ÏºÀ        2003/12/12 SystemBrain NullÎ
        .HWFBM1AX = fncNullCheck(rs("HWFBM1AX"))        ' ivealcP½ÏãÀ        2003/12/12 SystemBrain NullÎ
        .HWFBM1SH = rs("HWFBM1SH")                      ' ivealcPªèÊuQû
        .HWFBM1ST = rs("HWFBM1ST")                      ' ivealcPªèÊuQ_
        .HWFBM1SR = rs("HWFBM1SR")                      ' ivealcPªèÊuQÌ
        .HWFBM1HT = rs("HWFBM1HT")                      ' ivealcPÛØû@QÎ
        .HWFBM1HS = rs("HWFBM1HS")                      ' ivealcPÛØû@Q
        .HWFBM1NS = rs("HWFBM1NS")                      ' ivealcPM@
        .HWFBM1ET = fncNullCheck(rs("HWFBM1ET"))        ' ivealcPIðdsã      2003/12/12 SystemBrain NullÎ
        .HWFBM1SZ = rs("HWFBM1SZ")                      ' ivealcPªèð
        .HWFBM2AN = fncNullCheck(rs("HWFBM2AN"))        ' ivealcQ½ÏºÀ        2003/12/12 SystemBrain NullÎ
        .HWFBM2AX = fncNullCheck(rs("HWFBM2AX"))        ' ivealcQ½ÏãÀ        2003/12/12 SystemBrain NullÎ
        .HWFBM2SH = rs("HWFBM2SH")                      ' ivealcQªèÊuQû
        .HWFBM2ST = rs("HWFBM2ST")                      ' ivealcQªèÊuQ_
        .HWFBM2SR = rs("HWFBM2SR")                      ' ivealcQªèÊuQÌ
        .HWFBM2HT = rs("HWFBM2HT")                      ' ivealcQÛØû@QÎ
        .HWFBM2HS = rs("HWFBM2HS")                      ' ivealcQÛØû@Q
        .HWFBM2NS = rs("HWFBM2NS")                      ' ivealcQM@
        .HWFBM2ET = fncNullCheck(rs("HWFBM2ET"))        ' ivealcQIðdsã      2003/12/12 SystemBrain NullÎ
        .HWFBM2SZ = rs("HWFBM2SZ")                      ' ivealcQªèð
        .HWFBM3AN = fncNullCheck(rs("HWFBM3AN"))        ' ivealcR½ÏºÀ        2003/12/12 SystemBrain NullÎ
        .HWFBM3AX = fncNullCheck(rs("HWFBM3AX"))        ' ivealcR½ÏãÀ        2003/12/12 SystemBrain NullÎ
        .HWFBM3SH = rs("HWFBM3SH")                      ' ivealcRªèÊuQû
        .HWFBM3ST = rs("HWFBM3ST")                      ' ivealcRªèÊuQ_
        .HWFBM3SR = rs("HWFBM3SR")                      ' ivealcRªèÊuQÌ
        .HWFBM3HT = rs("HWFBM3HT")                      ' ivealcRÛØû@QÎ
        .HWFBM3HS = rs("HWFBM3HS")                      ' ivealcRÛØû@Q
        .HWFBM3NS = rs("HWFBM3NS")                      ' ivealcRM@
        .HWFBM3ET = fncNullCheck(rs("HWFBM3ET"))        ' ivealcRIðdsã      2003/12/12 SystemBrain NullÎ
        .HWFBM3SZ = rs("HWFBM3SZ")                      ' ivealcRªèð
        
        If Not IsNull(rs("HWFOSF1PTK")) Then .HWFOSF1PTK = rs("HWFOSF1PTK")   ' ivenrePp^æª
        If Not IsNull(rs("HWFOSF2PTK")) Then .HWFOSF2PTK = rs("HWFOSF2PTK")   ' ivenreQp^æª
        If Not IsNull(rs("HWFOSF3PTK")) Then .HWFOSF3PTK = rs("HWFOSF3PTK")   ' ivenreRp^æª
        If Not IsNull(rs("HWFOSF4PTK")) Then .HWFOSF4PTK = rs("HWFOSF4PTK")   ' ivenreSp^æª
        
'        If Not IsNull(rs("HWFBM1MBP")) Then .HWFBM1MBP = rs("HWFBM1MBP")      ' ivealcPÊàªz
'        If Not IsNull(rs("HWFBM2MBP")) Then .HWFBM2MBP = rs("HWFBM2MBP")      ' ivealcQÊàªz
'        If Not IsNull(rs("HWFBM3MBP")) Then .HWFBM3MBP = rs("HWFBM3MBP")      ' ivealcRÊàªz
        .HWFBM1MBP = fncNullCheck(rs("HWFBM1MBP"))      ' ivealcPÊàªz        2003/12/12 SystemBrain NullÎ
        .HWFBM2MBP = fncNullCheck(rs("HWFBM2MBP"))      ' ivealcQÊàªz        2003/12/12 SystemBrain NullÎ
        .HWFBM3MBP = fncNullCheck(rs("HWFBM3MBP"))      ' ivealcRÊàªz        2003/12/12 SystemBrain NullÎ
        If Not IsNull(rs("HWFBM1MCL")) Then .HWFBM1MCL = rs("HWFBM1MCL")      ' ivealcPÊàvZ
        If Not IsNull(rs("HWFBM2MCL")) Then .HWFBM2MCL = rs("HWFBM2MCL")      ' ivealcQÊàvZ
        If Not IsNull(rs("HWFBM3MCL")) Then .HWFBM3MCL = rs("HWFBM3MCL")      ' ivealcRÊàvZ
    
    '«ÏX M»fÇÁ 2006/02/15 SMPÎì ---------------
        'AN·x`FbNÌ×ÉTBCME025©çive`m·xðæ¾·é
        If Not IsNull(rs("HWFANTNP")) Then .HWFANTNP = rs("HWFANTNP")       ' ive`m·x
    'ªÏX M»fÇÁ 2006/02/15 SMPÎì ---------------
    
    End With
    Set rs = Nothing

    funGet_TBCME029 = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><<><><><
'Tv      :ïRèlðßéB
'Êß×Ò°À    :Ï¼         ,IO ,^        ,à¾
'          :CRYNUM        ,I  ,String    ,»Ô
'          :TopRs         ,I  ,Double    ,TOP¤è³ïRÀÑ
'          :TopPos        ,I  ,Double    ,TOP¤è³Êu
'          :BotRs         ,I  ,Double    ,TOP¤è³ïRÀÑ
'          :BotPos        ,I  ,Double    ,TOP¤è³Êu
'          :SuiPos        ,I  ,Double    ,èÊu
'          :Suitei  @    ,O  ,Double    ,èl
'          :ßèl        ,O  ,FUNCTION_RETURN,
'à¾      :»ÔATOP/BOTÌïRÀÑlAÊuæèïRèðs¤B
'ð      :2003/9/4 ì¬  }
Public Function new_ResSuitei(CRYNUM, TopRs, TOPPOS, BotRs, BOTPOS, SuiPos, Suitei As Double) As FUNCTION_RETURN
Dim cc As type_Coefficient  'ÀsÎÍvZp\¢Ì
Dim rp As type_ResPosCal    'èvZp\¢Ì
Dim Jikouhen As Double  'ÀsÎÍ
Dim wgtCharge As Long   '`[WÊ
Dim wgtTop As Double    'gbvdÊÀÑl
Dim wgtTopCut As Double 'gbvJbgdÊÀÑl
Dim DM As Double        '¼aP`RÌ½Ï
    
    new_ResSuitei = FUNCTION_RETURN_FAILURE
    
    ''ÀsÎÍpp[^æ¾ }`øãÎ QÆÖÏX 2008/04/23 SETsw Nakada
    If GetCoeffParams_new(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
'    If GetCoeffParams(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
        Debug.Print "ÎÍvZpp[^Ìæ¾É¸sµ½"
    End If
    
    ''ubNÌÀsÎÍðßé
    cc.DUNMENSEKI = AreaOfCircle(DM)    'fÊÏ
    cc.CHARGEWEIGHT = wgtCharge         '`[WÊ
    cc.TOPWEIGHT = wgtTop + wgtTopCut   'gbvdÊ
    cc.TOPSMPLPOS = TOPPOS
    cc.BOTSMPLPOS = BOTPOS
    cc.TOPRES = TopRs
    cc.BOTRES = BotRs
    
    Jikouhen = CoefficientCalculation(cc) 'ÀsÎÍvZ
    
    
    ''èïRlðßé
    If Jikouhen <> -9999 Then
        rp.COEFFICIENT = Jikouhen           'ÀsÎÍ
        rp.DUNMENSEKI = cc.DUNMENSEKI       'fÊÏ
        rp.CHARGEWEIGHT = cc.CHARGEWEIGHT   '`[WÊ
        rp.TOPWEIGHT = cc.TOPWEIGHT         'gbvdÊ
        rp.TOPSMPLPOS = TOPPOS
        rp.TOPRES = TopRs
        rp.target = SuiPos
        
        Suitei = ResCalculation(rp)         'èvZ
    Else
        new_ResSuitei = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    new_ResSuitei = FUNCTION_RETURN_SUCCESS

End Function
'
''><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><<><><><
''Tv      :ÎÍvZÉKvÈevdÊÀÑðæ¾·é
''Êß×Ò°À    :Ï¼        ,IO ,^        ,à¾
''          :CRYNUM        ,I  ,String    ,»Ô
''          :wgtCharge     ,O  ,Long      ,FàÊiñ`[WÊ|OñÜÅÌøã°dÊ|OñÜÅÌÄ¯Ìß¶¯ÄdÊj
''          :wgtTop        ,O  ,Double    ,gbvdÊÀÑl
''          :wgtTopCut     ,O  ,Double    ,gbvJbgdÊÀÑl
''          :DM            ,O  ,Double    ,¼aP`RÌ½Ï
''          :ßèl        ,O  ,FUNCTION_RETURN,
''à¾      :P{ø«AcÊø«É í¹ÄÀÑf[^ðæ¾·é
''ð      :2001/8/29 ì¬  ìº
'Public Function GetCoeffParams(ByVal CRYNUM$, wgtCharge As Long, wgtTop As Double, wgtTopCut As Double, DM As Double) As FUNCTION_RETURN
'Dim sql As String
'Dim rs As OraDynaset
'
'    On Error GoTo Err
'    GetCoeffParams = FUNCTION_RETURN_FAILURE
'    wgtCharge = 0
'    wgtTop = 0#
'    wgtTopCut = 0#
'    DM = 0#
'
'    sql = "select decode(RONAI,null,CHARGE,RONAI) as RONAI, WGHTTOP, WGTOPCUT, (DM1+DM2+DM3)/3.0 as DM " & _
'          "from TBCMH004 H004, " & _
'          "  (select sum(CHARGE) - sum(UPWEIGHT) - sum(WGTOPCUT) as RONAI" & _
'          "   From TBCMH004" & _
'          "   where (CRYNUM<'" & CRYNUM & "')" & _
'          "    and  (substr(CRYNUM,1,7)='" & Left$(CRYNUM, 7) & "')" & _
'          "  ) SUMDATA " & _
'          "where (CRYNUM='" & CRYNUM & "')"
'
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'    If rs.RecordCount > 0 Then
'        wgtCharge = rs("RONAI")
'        wgtTop = rs("WGHTTOP")
'        wgtTopCut = rs("WGTOPCUT")
'        DM = rs("DM")
'    End If
'    rs.Close
'
'    GetCoeffParams = FUNCTION_RETURN_SUCCESS
'
'proc_exit:
'    On Error GoTo 0
'    Exit Function
'
'Err:
'    Resume proc_exit
'End Function
'
''><><><><><><><><><><><><><><><><><><><><><>><><><><><><><><><><><><><><><><><><><><><><
''Tv      :ÊuÉÎ·éïRlðè·éB
''Êß×Ò°À    :Ï¼        ,IO ,^             ,à¾
''          :d             ,IO ,type_ResPosCal ,èvZ\¢Ì
''          :ßèl        ,O  ,Double         ,èïRl
''à¾      :
''ð      :2001/06/23@²ì MÆ@ì¬
'Public Function ResCalculation(d As type_ResPosCal) As Double
'    Dim GS As Double
'    Dim Ro As Double
'    Dim Gx As Double
'
'    On Error GoTo Err
'    GS = (d.DUNMENSEKI * HIJU_SILICONE * d.TOPSMPLPOS) / (d.CHARGEWEIGHT - d.TOPWEIGHT)
'    Ro = d.TOPRES * (1 - GS) ^ (d.COEFFICIENT - 1)
'    Gx = d.DUNMENSEKI * d.target * HIJU_SILICONE / (d.CHARGEWEIGHT - d.TOPWEIGHT)
'
'    ResCalculation = Ro / (1 - Gx) ^ (d.COEFFICIENT - 1)
'    On Error GoTo 0
'    Exit Function
'Err:
'    On Error GoTo 0
'    ResCalculation = -9999
'End Function

'------------------------------------------------
' TBCME050f[^æ¾
'------------------------------------------------

'Tv      :e[uuTBCME050v©çwèiÔÌR[hðo·éB
'Êß×Ò°À    :Ï¼        ,IO ,^                                   :à¾
'          :tHin          ,I  ,tFullHinban                          :iÔ
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :oR[h
'    @@  :sErrMsg @@  ,O  ,String     @@@@@@@@@@@    :G[bZ[W
'          :ßèl        ,O  ,FUNCTION_RETURN                      :oÌ¬Û
'à¾      :
'ð      :2006/08/15 VKì¬ Gsæs]¿ÇÁÎ SMP)kondoh

Public Function funGet_TBCME050(tHIN As tFullHinban, _
                                tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                                Optional sErrMsg As String = vbNullString) As FUNCTION_RETURN

    Dim sql         As String           'SQLSÌ
    Dim rs          As OraDynaset       'RecordSet
    Dim sDbName     As String

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "SB_GetSiyou.bas -- Function funGet_TBCME050"

    sDbName = "E050"
    'iEPBMD3½ÏºÀ(Oü),½ÏãÀ(Oü)ÇÁ@09/05/07 ooba
    sql = "SELECT hinban, mnorevno, factory, opecond, hepantnp"
    sql = sql & " ,hepof1ax ,hepof1mx ,hepof1et ,hepof1ns ,hepof1sz ,hepof1sh ,hepof1st ,hepof1sr ,hepof1ht ,hepof1hs "
    sql = sql & " ,hepof1km ,hepof1kn ,hepof1kh ,hepof1ku ,heposf1ptk"
    sql = sql & " ,hepof2ax ,hepof2mx ,hepof2et ,hepof2ns ,hepof2sz ,hepof2sh ,hepof2st ,hepof2sr ,hepof2ht ,hepof2hs"
    sql = sql & " ,hepof2km ,hepof2kn ,hepof2kh ,hepof2ku ,heposf2ptk"
    sql = sql & " ,hepof3ax ,hepof3mx ,hepof3et ,hepof3ns ,hepof3sz ,hepof3sh ,hepof3st ,hepof3sr ,hepof3ht ,hepof3hs"
    sql = sql & " ,hepof3km ,hepof3kn ,hepof3kh ,hepof3ku ,heposf3ptk"
    sql = sql & " ,hepbm1an ,hepbm1ax ,hepbm1et ,hepbm1ns ,hepbm1sz ,hepbm1sh ,hepbm1st ,hepbm1sr ,hepbm1ht ,hepbm1hs"
    sql = sql & " ,hepbm1km ,hepbm1kn ,hepbm1kh ,hepbm1ku ,hepbm1mbp ,hepbm1mcl"
    sql = sql & " ,hepbm2an ,hepbm2ax ,hepbm2et ,hepbm2ns ,hepbm2sz ,hepbm2sh ,hepbm2st ,hepbm2sr ,hepbm2ht ,hepbm2hs"
    sql = sql & " ,hepbm2km ,hepbm2kn ,hepbm2kh ,hepbm2ku ,hepbm2mbp ,hepbm2mcl"
    sql = sql & " ,hepbm3an ,hepbm3ax ,hepbm3gsan ,hepbm3gsax ,hepbm3et ,hepbm3ns ,hepbm3sz ,hepbm3sh ,hepbm3st ,hepbm3sr ,hepbm3ht ,hepbm3hs"
    sql = sql & " ,hepbm3km ,hepbm3kn ,hepbm3kh ,hepbm3ku ,hepbm3mbp ,hepbm3mcl"
    sql = sql & " FROM tbcme050 "
    sql = sql & " WHERE hinban = '" & tHIN.hinban & "' and "
    sql = sql & "      mnorevno = " & tHIN.mnorevno & " and "
    sql = sql & "      factory = '" & tHIN.factory & "' and "
    sql = sql & "      opecond = '" & tHIN.opecond & "'"

    ''f[^ðo·é
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDbName)
        funGet_TBCME050 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''oÊði[·é
     With tGetRec
        .HEPANTNP = fncNullCheck(rs("HEPANTNP"))                            ' iEPAN·x
        .HEPOF1AX = fncNullCheck(rs("HEPOF1AX"))                            ' iEPOSF1½ÏãÀ
        .HEPOF1MX = fncNullCheck(rs("HEPOF1MX"))                            ' iEPOSF1ãÀ
        .HEPOF1ET = fncNullCheck(rs("HEPOF1ET"))                            ' iEPOSF1IðETã
        .HEPOF1NS = IIf(IsNull(rs("HEPOF1NS")), "", rs("HEPOF1NS"))         ' iEPOSF1M@
        .HEPOF1SZ = IIf(IsNull(rs("HEPOF1SZ")), "", rs("HEPOF1SZ"))         ' iEPOSF1ªèð
        .HEPOF1SH = IIf(IsNull(rs("HEPOF1SH")), "", rs("HEPOF1SH"))         ' iEPOSF1ªèÊu_û
        .HEPOF1ST = IIf(IsNull(rs("HEPOF1ST")), "", rs("HEPOF1ST"))         ' iEPOSF1ªèÊu__
        .HEPOF1SR = IIf(IsNull(rs("HEPOF1SR")), "", rs("HEPOF1SR"))         ' iEPOSF1ªèÊu_Ì
        .HEPOF1HT = IIf(IsNull(rs("HEPOF1HT")), "", rs("HEPOF1HT"))         ' iEPOSF1ÛØû@_Î
        .HEPOF1HS = IIf(IsNull(rs("HEPOF1HS")), "", rs("HEPOF1HS"))         ' iEPOSF1ÛØû@_
        .HEPOF1KM = IIf(IsNull(rs("HEPOF1KM")), "", rs("HEPOF1KM"))         ' iEPOSF1¸px_
        .HEPOF1KN = IIf(IsNull(rs("HEPOF1KN")), "", rs("HEPOF1KN"))         ' iEPOSF1¸px_²
        .HEPOF1KH = IIf(IsNull(rs("HEPOF1KH")), "", rs("HEPOF1KH"))         ' iEPOSF1¸px_Û
        .HEPOF1KU = IIf(IsNull(rs("HEPOF1KU")), "", rs("HEPOF1KU"))         ' iEPOSF1¸px_³
        .HEPOSF1PTK = IIf(IsNull(rs("HEPOSF1PTK")), "", rs("HEPOSF1PTK"))   ' iEPOSF1ÊßÀÝæª
        .HEPOF2AX = fncNullCheck(rs("HEPOF2AX"))                            ' iEPOSF2½ÏãÀ
        .HEPOF2MX = fncNullCheck(rs("HEPOF2MX"))                            ' iEPOSF2ãÀ
        .HEPOF2ET = fncNullCheck(rs("HEPOF2ET"))                            ' iEPOSF2IðETã
        .HEPOF2NS = IIf(IsNull(rs("HEPOF2NS")), "", rs("HEPOF2NS"))         ' iEPOSF2M@
        .HEPOF2SZ = IIf(IsNull(rs("HEPOF2SZ")), "", rs("HEPOF2SZ"))         ' iEPOSF2ªèð
        .HEPOF2SH = IIf(IsNull(rs("HEPOF2SH")), "", rs("HEPOF2SH"))         ' iEPOSF2ªèÊu_û
        .HEPOF2ST = IIf(IsNull(rs("HEPOF2ST")), "", rs("HEPOF2ST"))         ' iEPOSF2ªèÊu__
        .HEPOF2SR = IIf(IsNull(rs("HEPOF2SR")), "", rs("HEPOF2SR"))         ' iEPOSF2ªèÊu_Ì
        .HEPOF2HT = IIf(IsNull(rs("HEPOF2HT")), "", rs("HEPOF2HT"))         ' iEPOSF2ÛØû@_Î
        .HEPOF2HS = IIf(IsNull(rs("HEPOF2HS")), "", rs("HEPOF2HS"))         ' iEPOSF2ÛØû@_
        .HEPOF2KM = IIf(IsNull(rs("HEPOF2KM")), "", rs("HEPOF2KM"))         ' iEPOSF2¸px_
        .HEPOF2KN = IIf(IsNull(rs("HEPOF2KN")), "", rs("HEPOF2KN"))         ' iEPOSF2¸px_²
        .HEPOF2KH = IIf(IsNull(rs("HEPOF2KH")), "", rs("HEPOF2KH"))         ' iEPOSF2¸px_Û
        .HEPOF2KU = IIf(IsNull(rs("HEPOF2KU")), "", rs("HEPOF2KU"))         ' iEPOSF2¸px_³
        .HEPOSF2PTK = IIf(IsNull(rs("HEPOSF2PTK")), "", rs("HEPOSF2PTK"))   ' iEPOSF2ÊßÀÝæª
        .HEPOF3AX = fncNullCheck(rs("HEPOF3AX"))                            ' iEPOSF3½ÏãÀ
        .HEPOF3MX = fncNullCheck(rs("HEPOF3MX"))                            ' iEPOSF3ãÀ
        .HEPOF3ET = fncNullCheck(rs("HEPOF3ET"))                            ' iEPOSF3IðETã
        .HEPOF3NS = IIf(IsNull(rs("HEPOF3NS")), "", rs("HEPOF3NS"))         ' iEPOSF3M@
        .HEPOF3SZ = IIf(IsNull(rs("HEPOF3SZ")), "", rs("HEPOF3SZ"))         ' iEPOSF3ªèð
        .HEPOF3SH = IIf(IsNull(rs("HEPOF3SH")), "", rs("HEPOF3SH"))         ' iEPOSF3ªèÊu_û
        .HEPOF3ST = IIf(IsNull(rs("HEPOF3ST")), "", rs("HEPOF3ST"))         ' iEPOSF3ªèÊu__
        .HEPOF3SR = IIf(IsNull(rs("HEPOF3SR")), "", rs("HEPOF3SR"))         ' iEPOSF3ªèÊu_Ì
        .HEPOF3HT = IIf(IsNull(rs("HEPOF3HT")), "", rs("HEPOF3HT"))         ' iEPOSF3ÛØû@_Î
        .HEPOF3HS = IIf(IsNull(rs("HEPOF3HS")), "", rs("HEPOF3HS"))         ' iEPOSF3ÛØû@_
        .HEPOF3KM = IIf(IsNull(rs("HEPOF3KM")), "", rs("HEPOF3KM"))         ' iEPOSF3¸px_
        .HEPOF3KN = IIf(IsNull(rs("HEPOF3KN")), "", rs("HEPOF3KN"))         ' iEPOSF3¸px_²
        .HEPOF3KH = IIf(IsNull(rs("HEPOF3KH")), "", rs("HEPOF3KH"))         ' iEPOSF3¸px_Û
        .HEPOF3KU = IIf(IsNull(rs("HEPOF3KU")), "", rs("HEPOF3KU"))         ' iEPOSF3¸px_³
        .HEPOSF3PTK = IIf(IsNull(rs("HEPOSF3PTK")), "", rs("HEPOSF3PTK"))   ' iEPOSF3ÊßÀÝæª
        .HEPBM1AN = fncNullCheck(rs("HEPBM1AN"))                            ' iEPBMD1½ÏºÀ
        .HEPBM1AX = fncNullCheck(rs("HEPBM1AX"))                            ' iEPBMD1½ÏãÀ
        .HEPBM1ET = fncNullCheck(rs("HEPBM1ET"))                            ' iEPBMD1IðETã
        .HEPBM1NS = IIf(IsNull(rs("HEPBM1NS")), "", rs("HEPBM1NS"))         ' iEPBMD1M@
        .HEPBM1SZ = IIf(IsNull(rs("HEPBM1SZ")), "", rs("HEPBM1SZ"))         ' iEPBMD1ªèð
        .HEPBM1SH = IIf(IsNull(rs("HEPBM1SH")), "", rs("HEPBM1SH"))         ' iEPBMD1ªèÊu_û
        .HEPBM1ST = IIf(IsNull(rs("HEPBM1ST")), "", rs("HEPBM1ST"))         ' iEPBMD1ªèÊu__
        .HEPBM1SR = IIf(IsNull(rs("HEPBM1SR")), "", rs("HEPBM1SR"))         ' iEPBMD1ªèÊu_Ì
        .HEPBM1HT = IIf(IsNull(rs("HEPBM1HT")), "", rs("HEPBM1HT"))         ' iEPBMD1ÛØû@_Î
        .HEPBM1HS = IIf(IsNull(rs("HEPBM1HS")), "", rs("HEPBM1HS"))         ' iEPBMD1ÛØû@_
        .HEPBM1KM = IIf(IsNull(rs("HEPBM1KM")), "", rs("HEPBM1KM"))         ' iEPBMD1¸px_
        .HEPBM1KN = IIf(IsNull(rs("HEPBM1KN")), "", rs("HEPBM1KN"))         ' iEPBMD1¸px_²
        .HEPBM1KH = IIf(IsNull(rs("HEPBM1KH")), "", rs("HEPBM1KH"))         ' iEPBMD1¸px_Û
        .HEPBM1KU = IIf(IsNull(rs("HEPBM1KU")), "", rs("HEPBM1KU"))         ' iEPBMD1¸px_³
        .HEPBM1MBP = fncNullCheck(rs("HEPBM1MBP"))                          ' iEPBMD1Êàªz
        .HEPBM1MCL = IIf(IsNull(rs("HEPBM1MCL")), "", rs("HEPBM1MCL"))      ' iEPBMD1ÊàvZ
        .HEPBM2AN = fncNullCheck(rs("HEPBM2AN"))                            ' iEPBMD2½ÏºÀ
        .HEPBM2AX = fncNullCheck(rs("HEPBM2AX"))                            ' iEPBMD2½ÏãÀ
        .HEPBM2ET = fncNullCheck(rs("HEPBM2ET"))                            ' iEPBMD2IðETã
        .HEPBM2NS = IIf(IsNull(rs("HEPBM2NS")), "", rs("HEPBM2NS"))         ' iEPBMD2M@
        .HEPBM2SZ = IIf(IsNull(rs("HEPBM2SZ")), "", rs("HEPBM2SZ"))         ' iEPBMD2ªèð
        .HEPBM2SH = IIf(IsNull(rs("HEPBM2SH")), "", rs("HEPBM2SH"))         ' iEPBMD2ªèÊu_û
        .HEPBM2ST = IIf(IsNull(rs("HEPBM2ST")), "", rs("HEPBM2ST"))         ' iEPBMD2ªèÊu__
        .HEPBM2SR = IIf(IsNull(rs("HEPBM2SR")), "", rs("HEPBM2SR"))         ' iEPBMD2ªèÊu_Ì
        .HEPBM2HT = IIf(IsNull(rs("HEPBM2HT")), "", rs("HEPBM2HT"))         ' iEPBMD2ÛØû@_Î
        .HEPBM2HS = IIf(IsNull(rs("HEPBM2HS")), "", rs("HEPBM2HS"))         ' iEPBMD2ÛØû@_
        .HEPBM2KM = IIf(IsNull(rs("HEPBM2KM")), "", rs("HEPBM2KM"))         ' iEPBMD2¸px_
        .HEPBM2KN = IIf(IsNull(rs("HEPBM2KN")), "", rs("HEPBM2KN"))         ' iEPBMD2¸px_²
        .HEPBM2KH = IIf(IsNull(rs("HEPBM2KH")), "", rs("HEPBM2KH"))         ' iEPBMD2¸px_Û
        .HEPBM2KU = IIf(IsNull(rs("HEPBM2KU")), "", rs("HEPBM2KU"))         ' iEPBMD2¸px_³
        .HEPBM2MBP = fncNullCheck(rs("HEPBM2MBP"))                          ' iEPBMD2Êàªz
        .HEPBM2MCL = IIf(IsNull(rs("HEPBM2MCL")), "", rs("HEPBM2MCL"))      ' iEPBMD2ÊàvZ
        .HEPBM3AN = fncNullCheck(rs("HEPBM3AN"))                            ' iEPBMD3½ÏºÀ
        .HEPBM3AX = fncNullCheck(rs("HEPBM3AX"))                            ' iEPBMD3½ÏãÀ
        .HEPBM3GSAN = fncNullCheck(rs("HEPBM3GSAN"))                        ' iEPBMD3½ÏºÀ(Oü)@09/05/07 ooba
        .HEPBM3GSAX = fncNullCheck(rs("HEPBM3GSAX"))                        ' iEPBMD3½ÏãÀ(Oü)@09/05/07 ooba
        .HEPBM3ET = fncNullCheck(rs("HEPBM3ET"))                            ' iEPBMD3IðETã
        .HEPBM3NS = IIf(IsNull(rs("HEPBM3NS")), "", rs("HEPBM3NS"))         ' iEPBMD3M@
        .HEPBM3SZ = IIf(IsNull(rs("HEPBM3SZ")), "", rs("HEPBM3SZ"))         ' iEPBMD3ªèð
        .HEPBM3SH = IIf(IsNull(rs("HEPBM3SH")), "", rs("HEPBM3SH"))         ' iEPBMD3ªèÊu_û
        .HEPBM3ST = IIf(IsNull(rs("HEPBM3ST")), "", rs("HEPBM3ST"))         ' iEPBMD3ªèÊu__
        .HEPBM3SR = IIf(IsNull(rs("HEPBM3SR")), "", rs("HEPBM3SR"))         ' iEPBMD3ªèÊu_Ì
        .HEPBM3HT = IIf(IsNull(rs("HEPBM3HT")), "", rs("HEPBM3HT"))         ' iEPBMD3ÛØû@_Î
        .HEPBM3HS = IIf(IsNull(rs("HEPBM3HS")), "", rs("HEPBM3HS"))         ' iEPBMD3ÛØû@_
        .HEPBM3KM = IIf(IsNull(rs("HEPBM3KM")), "", rs("HEPBM3KM"))         ' iEPBMD3¸px_
        .HEPBM3KN = IIf(IsNull(rs("HEPBM3KN")), "", rs("HEPBM3KN"))         ' iEPBMD3¸px_²
        .HEPBM3KH = IIf(IsNull(rs("HEPBM3KH")), "", rs("HEPBM3KH"))         ' iEPBMD3¸px_Û
        .HEPBM3KU = IIf(IsNull(rs("HEPBM3KU")), "", rs("HEPBM3KU"))         ' iEPBMD3¸px_³
        .HEPBM3MBP = fncNullCheck(rs("HEPBM3MBP"))                          ' iEPBMD3Êàªz
        .HEPBM3MCL = IIf(IsNull(rs("HEPBM3MCL")), "", rs("HEPBM3MCL"))      ' iEPBMD3ÊàvZ
    End With
    Set rs = Nothing

    funGet_TBCME050 = FUNCTION_RETURN_SUCCESS
  
proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

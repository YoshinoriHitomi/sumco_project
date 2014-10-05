Attribute VB_Name = "s_cmbc055_SQL"
Option Explicit
'
'================================================
' DBANZXÖ
' è`àe: TBCMB019 (FRSZ³îñ)
' QÆ@@: 060200_Se[u
'================================================

'------------------------------------------------
' [Uè`^Ìé¾
'------------------------------------------------
Public Type typ_cmjc001j_Disp
    GOUKI       As String * 3       ' @
    INPDATE     As Date             ' út
    FTIROIL     As Double           ' FTIRiOiá)
    FTIROIM     As Double           ' FTIRiOij
    FTIROIH     As Double           ' FTIRiOij
    MS1OIL      As Double           ' ªèTv1iOiá)
    MS1OIM      As Double           ' ªèTv1iOij
    MS1OIH      As Double           ' ªèTv1iOij
    MS2OIL      As Double           ' ªèTv2iOiá)
    MS2OIM      As Double           ' ªèTv2iOij
    MS2OIH      As Double           ' ªèTv2iOij
    MS3OIL      As Double           ' ªèTv3iOiá)
    MS3OIM      As Double           ' ªèTv3iOij
    MS3OIH      As Double           ' ªèTv3iOij
    MS4OIL      As Double           ' ªèTv4iOiá)
    MS4OIM      As Double           ' ªèTv4iOij
    MS4OIH      As Double           ' ªèTv4iOij
    MS5OIL      As Double           ' ªèTv5iOiá)
    MS5OIM      As Double           ' ªèTv5iOij
    MS5OIH      As Double           ' ªèTv5iOij
    MSAVEOIL    As Double           ' ªè½ÏiOiá)
    MSAVEOIM    As Double           ' ªè½ÏiOij
    MSAVEOIH    As Double           ' ªè½ÏiOij
    MSSGOIL     As Double           ' ªèÐiOiá)
    MSSGOIM     As Double           ' ªèÐiOij
    MSSGOIH     As Double           ' ªèÐiOij
    MSPSGOIL    As Double           ' ªèAVE+ÐiOiá)
    MSPSGOIM    As Double           ' ªèAVE+ÐiOij
    MSPSGOIH    As Double           ' ªèAVE+ÐiOij
    MSNSGOIL    As Double           ' ªèAVE-ÐiOiá)
    MSNSGOIM    As Double           ' ªèAVE-ÐiOij
    MSNSGOIH    As Double           ' ªèAVE-ÐiOij
    MINOIL      As Double           ' MINiOiá)
    MINOIM      As Double           ' MINiOij
    MINOIH      As Double           ' MINiOij
    MAXOIL      As Double           ' MAXiOiá)
    MAXOIM      As Double           ' MAXiOij
    MAXOIH      As Double           ' MAXiOij
    SGCK1OIL    As Double           ' ÐckTv1iOiá)
    SGCK1OIM    As Double           ' ÐckTv1iOij
    SGCK1OIH    As Double           ' ÐckTv1iOij
    SGCK2OIL    As Double           ' ÐckTv2iOiá)
    SGCK2OIM    As Double           ' ÐckTv2iOij
    SGCK2OIH    As Double           ' ÐckTv2iOij
    SGCK3OIL    As Double           ' ÐckTv3iOiá)
    SGCK3OIM    As Double           ' ÐckTv3iOij
    SGCK3OIH    As Double           ' ÐckTv3iOij
    SGCK4OIL    As Double           ' ÐckTv4iOiá)
    SGCK4OIM    As Double           ' ÐckTv4iOij
    SGCK4OIH    As Double           ' ÐckTv4iOij
    SGCK5OIL    As Double           ' ÐckTv5iOiá)
    SGCK5OIM    As Double           ' ÐckTv5iOij
    SGCK5OIH    As Double           ' ÐckTv5iOij
    SGCKDOIL    As Double           ' Ðckf[^iOiá)
    SGCKDOIM    As Double           ' Ðckf[^iOij
    SGCKDOIH    As Double           ' Ðckf[^iOij
    SGCKAOIL    As Double           ' Ðck½ÏiOiá)
    SGCKAAOIM   As Double           ' Ðck½ÏiOij
    SGCKAOIH    As Double           ' Ðck½ÏiOij
    SGNOIL      As Double           ' ÐckÐiOiá)
    SGNOIM      As Double           ' ÐckÐiOij
    SGNOIH      As Double           ' ÐckÐiOij
    FTIRKOIL    As Double           ' FTIR·ZiOiá)
    FTIRKOIM    As Double           ' FTIR·ZiOij
    FTIRKOIH    As Double           ' FTIR·ZiOij
    EFFECTTM    As Integer          ' LøÔ
    YCOEF       As Double           ' eshq·Z®ixØÐj
    XCOEF       As Double           ' eshq·Z®iwWj
    RSQUARE     As Double           ' qQæ
    SGCKST      As Double           ' Ð»èî
    SGCKOIL     As String * 1       ' Ð»èiOiá)
    SGCKOIM     As String * 1       ' Ð»èiOij
    SGCKOIH     As String * 1       ' Ð»èiOij
    FTIRCKST    As Double           ' FTIR·Z»èî
    FTIRCKOIL   As String * 1       ' FTIR·Z»èiOiá)
    FTIRCKOIM   As String * 1       ' FTIR·Z»èiOij
    FTIRCKOIH   As String * 1       ' FTIR·Z»èiOij
    MS6OIL      As Double           ' ªèTv6iOiá)
    MS6OIM      As Double           ' ªèTv6iOij
    MS6OIH      As Double           ' ªèTv6iOij
    SGCK6OIL    As Double           ' ÐckTv6iOiá)
    SGCK6OIM    As Double           ' ÐckTv6iOij
    SGCK6OIH    As Double           ' ÐckTv6iOij
    CVOIL       As Double           ' CV(%)iOiá)
    CVOIM       As Double           ' CV(%)iOij
    CVOIH       As Double           ' CV(%)iOij
  '  TSTAFFID As String * 8          ' o^ÐõID
  '  REGDATE As Date                 ' o^út
  '  KSTAFFID As String * 8          ' XVÐõID
  '  UPDDATE As Date                 ' XVút
  '  SENDFLAG As String * 1          ' MtO
  '  SENDDATE As Date                ' Mút
End Type

'''''------------------------------------------------
''''' DBANZXÖ
'''''------------------------------------------------
''''
'''''Tv      :e[uuTBCMB019v©çðÉ Á½R[hðo·é
'''''Êß×Ò°À    :Ï¼        ,IO ,^           ,à¾
'''''          :record        ,O  ,typ_cmjc001j_Disp ,oR[h
'''''          :GOUK          ,I  ,String       ,u@v(SQLÌoð)
'''''          :ßèl        ,O  ,FUNCTION_RETURN ,oÌ¬Û
'''''à¾      :u@v=øÅA©ÂuútvªÅVÌf[^ðo·é
'''''ð      :2001/06/20ì¬@·ì
''''Public Function DBDRV_Getcmjc001j_Disp(record As typ_cmjc001j_Disp, GOUK$) As FUNCTION_RETURN
''''Dim sql As String       'SQLSÌ
''''Dim sqlBase As String   'SQLî{(WHEREßÌOÜÅ)
''''Dim sqlWhere As String  'SQLÌWHEREª
''''Dim sqlGroup As String  'SQLÌGROUPª
''''Dim rs As OraDynaset    'RecordSet
''''Dim recCnt As Long      'R[h
''''Dim i As Long
''''
''''    DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_FAILURE
''''
''''    ''SQLðgÝ§Äé
''''
''''    'G[nhÌÝè
''''    On Error GoTo proc_err
''''    gErr.Push "s_cmzcF_cmjc001j_SQL.bas -- Function DBDRV_Getcmjc001j_Disp"
''''
''''    sqlBase = "Select GOUKI, MAX(INPDATE) ""INPDATE"", FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
''''              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
''''              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
''''              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
''''              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
''''              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE "
''''    sqlBase = sqlBase & "From TBCMB019"
''''    ''oð(»ÝÌßÙNO)Ìæèoµ
''''    sqlWhere = "WHERE(GOUKI=" & GOUK & ") "
''''    sqlGroup = "GROUP BY GOUKI"
''''    sql = sqlBase & sqlWhere & sqlGroup
''''
''''    ''f[^ðo·é
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''    If rs Is Nothing Then
''''        ReDim records(0)
''''        DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_FAILURE
''''        GoTo proc_exit
''''    End If
''''
''''    ''oÊði[·é
''''    With record
''''        .GOUKI = rs("GOUKI")             ' @
''''        .INPDATE = rs("INPDATE")         ' út
''''        .FTIRFZI = rs("FTIRFZI")         ' FTIRiFZ)
''''        .FTIRCZH = rs("FTIRCZH")         ' FTIRiCZj
''''        .FTIRCZC = rs("FTIRCZC")         ' FTIRiCZj
''''        .MS1FZ = rs("MS1FZ")             ' ªèTv1iFZ)
''''        .MS1CZ1 = rs("MS1CZ1")           ' ªèTv1iCZ-1)
''''        .MS1CZ2 = rs("MS1CZ2")           ' ªèTv1iCZ-2)
''''        .MS2FZ = rs("MS2FZ")             ' ªèTv2iFZ)
''''        .MS2CZ1 = rs("MS2CZ1")           ' ªèTv2iCZ-1)
''''        .MS2CZ2 = rs("MS2CZ2")           ' ªèTv2iCZ-2)
''''        .MS3FZ = rs("MS3FZ")             ' ªèTv3iFZ)
''''        .MS3CZ1 = rs("MS3CZ1")           ' ªèTv3iCZ-1)
''''        .MS3CZ2 = rs("MS3CZ2")           ' ªèTv3iCZ-2)
''''        .MS4FZ = rs("MS4FZ")             ' ªèTv4iFZ)
''''        .MS4CZ1 = rs("MS4CZ1")           ' ªèTv4iCZ-1)
''''        .MS4CZ2 = rs("MS4CZ2")           ' ªèTv4iCZ-2)
''''        .MS5FZ = rs("MS5FZ")             ' ªèTv5iFZ)
''''        .MS5CZ1 = rs("MS5CZ1")           ' ªèTv5iCZ-1)
''''        .MS5CZ2 = rs("MS5CZ2")           ' ªèTv5iCZ-2)
''''        .MSAVEFZ = rs("MSAVEFZ")         ' ªè½ÏiFZj
''''        .MSAVECZ1 = rs("MSAVECZ1")       ' ªè½ÏiCZ-1j
''''        .MSAVECZ2 = rs("MSAVECZ2")       ' ªè½ÏiCZ-2j
''''        .MSSGFZ = rs("MSSGFZ")           ' ªèÐiFZj
''''        .MSSGCZ1 = rs("MSSGCZ1")         ' ªèÐiCZ-1j
''''        .MSSGCZ2 = rs("MSSGCZ2")         ' ªèÐiCZ-2j
''''        .MSPSGFZ = rs("MSPSGFZ")         ' ªèAVE+ÐiFZj
''''        .MSPSGCZ1 = rs("MSPSGCZ1")       ' ªèAVE+ÐiCZ-1j
''''        .MSPSGCZ2 = rs("MSPSGCZ2")       ' ªèAVE+ÐiCZ-2j
''''        .MSNSGFZ = rs("MSNSGFZ")         ' ªèAVE-ÐiFZj
''''        .MSNSGCZ1 = rs("MSNSGCZ1")       ' ªèAVE-ÐiCZ-1j
''''        .MSNSGCZ2 = rs("MSNSGCZ2")       ' ªèAVE-ÐiCZ-2j
''''        .MINFZ = rs("MINFZ")             ' MINiFZj
''''        .MINCZ1 = rs("MINCZ1")           ' MINiCZ-1j
''''        .MINCZ2 = rs("MINCZ2")           ' MINiCZ-2j
''''        .MAXFZ = rs("MAXFZ")             ' MAXiFZj
''''        .MAXCZ1 = rs("MAXCZ1")           ' MAXiCZ-1j
''''        .MAXCZ2 = rs("MAXCZ2")           ' MAXiCZ-2j
''''        .SGCK1FZ = rs("SGCK1FZ")         ' ÐckTv1iFZ)
''''        .SGCK1CZ1 = rs("SGCK1CZ1")       ' ÐckTv1iCZ-1)
''''        .SGCK1CZ2 = rs("SGCK1CZ2")       ' ÐckTv1iCZ-2)
''''        .SGCK2FZ = rs("SGCK2FZ")         ' ÐckTv2iFZ)
''''        .SGCK2CZ1 = rs("SGCK2CZ1")       ' ÐckTv2iCZ-1)
''''        .SGCK2CZ2 = rs("SGCK2CZ2")       ' ÐckTv2iCZ-2)
''''        .SGCK3FZ = rs("SGCK3FZ")         ' ÐckTv3iFZ)
''''        .SGCK3CZ1 = rs("SGCK3CZ1")       ' ÐckTv3iCZ-1)
''''        .SGCK3CZ2 = rs("SGCK3CZ2")       ' ÐckTv3iCZ-2)
''''        .SGCK4FZ = rs("SGCK4FZ")         ' ÐckTv4iFZ)
''''        .SGCK4CZ1 = rs("SGCK4CZ1")       ' ÐckTv4iCZ-1)
''''        .SGCK4CZ2 = rs("SGCK4CZ2")       ' ÐckTv4iCZ-2)
''''        .SGCK5FZ = rs("SGCK5FZ")         ' ÐckTv5iFZ)
''''        .SGCK5CZ1 = rs("SGCK5CZ1")       ' ÐckTv5iCZ-1)
''''        .SGCK5CZ2 = rs("SGCK5CZ2")       ' ÐckTv5iCZ-2)
''''        .SGCKDFZ = rs("SGCKDFZ")         ' Ðckf[^iFZj
''''        .SGCKDCZ1 = rs("SGCKDCZ1")       ' Ðckf[^iCZ-1j
''''        .SGCKDCZ2 = rs("SGCKDCZ2")       ' Ðckf[^iCZ-2j
''''        .SGCKAFZ = rs("SGCKAFZ")         ' Ðck½ÏiFZj
''''        .SGCKAACZ1 = rs("SGCKAACZ1")     ' Ðck½ÏiCZ-1j
''''        .SGCKACZ2 = rs("SGCKACZ2")       ' Ðck½ÏiCZ-2j
''''        .SGNFZ = rs("SGNFZ")             ' ÐckÐiFZj
''''        .SGNCZ1 = rs("SGNCZ1")           ' ÐckÐ CZ-1j
''''        .SGNCZ2 = rs("SGNCZ2")           ' ÐckÐiCZ-2j
''''        .FTIRFZ = rs("FTIRFZ")           ' FTIR·ZiFZj
''''        .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR·ZiCZ-1j
''''        .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR·ZiCZ-2j
''''        .EFFECTTM = rs("EFFECTTM")       ' LøÔ
''''        .YCOEF = rs("YCOEF")             ' eshq·Z®ixØÐj
''''        .XCOEF = rs("XCOEF")             ' eshq·Z®iwWj
''''        .RSQUARE = rs("RSQUARE")         ' qQæ
''''    End With
''''    rs.Close
''''
''''    DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_SUCCESS
''''
''''proc_exit:
''''    'I¹
''''    gErr.Pop
''''    Exit Function
''''
''''proc_err:
''''    'G[nh
''''    Debug.Print "====== Error SQL ======"
''''    Debug.Print sql
''''    gErr.HandleError
''''    Resume proc_exit
''''End Function


'------------------------------------------------
' DBANZXÖ
'------------------------------------------------

'Tv      :øÅn³ê½R[hðTBCMB019ÉÇÁ·é
'Êß×Ò°À    :Ï¼        ,IO ,^            ,à¾
'          :record        ,I  ,typ_cmjc001j_Disp ,oR[h
'          :TSTAFFID      ,I  ,String       ,o^ÐõID
'          :ßèl        ,O  ,FUNCTION_RETURN ,oÌ¬Û
'à¾      :
'ð      :
Public Function DBDRV_Getcmjc001j_Exec(record As typ_cmjc001j_Disp, TSTAFFID$) As FUNCTION_RETURN
    Dim sql As String           'SQLSÌ
    Dim SetDate  As Variant     'üÍút

    DBDRV_Getcmjc001j_Exec = FUNCTION_RETURN_FAILURE
    
    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001j_SQL.bas -- Function DBDRV_Getcmjc001j_Exec"

    SetDate = Format$(record.INPDATE, "yyyy-mm-dd hh:mm:ss")
  
    ''SQLðgÝ§Äé
    sql = "Insert into TBCMB019 ("
    sql = sql & "  GOUKI"                   '' @
    sql = sql & ", INPDATE"                 '' út
    sql = sql & ", FTIROIL"                 '' FTIRiOiá)
    sql = sql & ", FTIROIM"                 '' FTIRiOi)
    sql = sql & ", FTIROIH"                 '' FTIRiOi)
    sql = sql & ", MS1OIL"                  '' ªèTv1iOiá)
    sql = sql & ", MS1OIM"                  '' ªèTv1iOi)
    sql = sql & ", MS1OIH"                  '' ªèTv1iOi)
    sql = sql & ", MS2OIL"                  '' ªèTv2iOiá)
    sql = sql & ", MS2OIM"                  '' ªèTv2iOi)
    sql = sql & ", MS2OIH"                  '' ªèTv2iOi)
    sql = sql & ", MS3OIL"                  '' ªèTv3iOiá)
    sql = sql & ", MS3OIM"                  '' ªèTv3iOi)
    sql = sql & ", MS3OIH"                  '' ªèTv3iOi)
    sql = sql & ", MS4OIL"                  '' ªèTv4iOiá)
    sql = sql & ", MS4OIM"                  '' ªèTv4iOi)
    sql = sql & ", MS4OIH"                  '' ªèTv4iOi)
    sql = sql & ", MS5OIL"                  '' ªèTv5iOiá)
    sql = sql & ", MS5OIM"                  '' ªèTv5iOi)
    sql = sql & ", MS5OIH"                  '' ªèTv5iOi)
    sql = sql & ", MSAVEOIL"                '' ªè½ÏiOiá)
    sql = sql & ", MSAVEOIM"                '' ªè½ÏiOi)
    sql = sql & ", MSAVEOIH"                '' ªè½ÏiOi)
    sql = sql & ", MSSGOIL"                 '' ªèÐiOiá)
    sql = sql & ", MSSGOIM"                 '' ªèÐiOi)
    sql = sql & ", MSSGOIH"                 '' ªèÐiOi)
    sql = sql & ", MSPSGOIL"                '' ªèAVE+ÐiOiá)
    sql = sql & ", MSPSGOIM"                '' ªèAVE+ÐiOi)
    sql = sql & ", MSPSGOIH"                '' ªèAVE+ÐiOi)
    sql = sql & ", MSNSGOIL"                '' ªèAVE-ÐiOiá)
    sql = sql & ", MSNSGOIM"                '' ªèAVE-ÐiOi)
    sql = sql & ", MSNSGOIH"                '' ªèAVE-ÐiOi)
    sql = sql & ", MINOIL"                  '' MINiOiá)
    sql = sql & ", MINOIM"                  '' MINiOi)
    sql = sql & ", MINOIH"                  '' MINiOi)
    sql = sql & ", MAXOIL"                  '' MAXiOiá)
    sql = sql & ", MAXOIM"                  '' MAXiOi)
    sql = sql & ", MAXOIH"                  '' MAXiOi)
    sql = sql & ", SGCK1OIL"                '' ÐckTv1iOiá)
    sql = sql & ", SGCK1OIM"                '' ÐckTv1iOi)
    sql = sql & ", SGCK1OIH"                '' ÐckTv1iOi)
    sql = sql & ", SGCK2OIL"                '' ÐckTv2iOiá)
    sql = sql & ", SGCK2OIM"                '' ÐckTv2iOi)
    sql = sql & ", SGCK2OIH"                '' ÐckTv2iOi)
    sql = sql & ", SGCK3OIL"                '' ÐckTv3iOiá)
    sql = sql & ", SGCK3OIM"                '' ÐckTv3iOi)
    sql = sql & ", SGCK3OIH"                '' ÐckTv3iOi)
    sql = sql & ", SGCK4OIL"                '' ÐckTv4iOiá)
    sql = sql & ", SGCK4OIM"                '' ÐckTv4iOi)
    sql = sql & ", SGCK4OIH"                '' ÐckTv4iOi)
    sql = sql & ", SGCK5OIL"                '' ÐckTv5iOiá)
    sql = sql & ", SGCK5OIM"                '' ÐckTv5iOi)
    sql = sql & ", SGCK5OIH"                '' ÐckTv5iOi)
    sql = sql & ", SGCKDOIL"                '' Ðckf[^iOiá)
    sql = sql & ", SGCKDOIM"                '' Ðckf[^iOi)
    sql = sql & ", SGCKDOIH"                '' Ðckf[^iOi)
    sql = sql & ", SGCKAOIL"                '' Ðck½ÏiOiá)
    sql = sql & ", SGCKAAOIM"               '' Ðck½ÏiOi)
    sql = sql & ", SGCKAOIH"                '' Ðck½ÏiOi)
    sql = sql & ", SGNOIL"                  '' ÐckÐiOiá)
    sql = sql & ", SGNOIM"                  '' ÐckÐiOi)
    sql = sql & ", SGNOIH"                  '' ÐckÐiOi)
    sql = sql & ", FTIRKOIL"                '' FTIR·ZiOiá)
    sql = sql & ", FTIRKOIM"                '' FTIR·ZiOi)
    sql = sql & ", FTIRKOIH"                '' FTIR·ZiOi)
    sql = sql & ", EFFECTTM"                '' LøÔ
    sql = sql & ", YCOEF"                   '' eshq·Z®ixØÐj
    sql = sql & ", XCOEF"                   '' eshq·Z®iwWj
    sql = sql & ", RSQUARE"                 '' qQæ
    sql = sql & ", SGCKST"                  '' Ð»èî
    sql = sql & ", SGCKOIL"                 '' Ð»èiOiá)
    sql = sql & ", SGCKOIM"                 '' Ð»èiOi)
    sql = sql & ", SGCKOIH"                 '' Ð»èiOi)
    sql = sql & ", FTIRCKST"                '' FTIR·Z»èî
    sql = sql & ", FTIRCKOIL"               '' FTIR·Z»èiOiá)
    sql = sql & ", FTIRCKOIM"               '' FTIR·Z»èiOi)
    sql = sql & ", FTIRCKOIH"               '' FTIR·Z»èiOi)
    sql = sql & ", MS6OIL"                  '' ªèTv6iOiá)
    sql = sql & ", MS6OIM"                  '' ªèTv6iOi)
    sql = sql & ", MS6OIH"                  '' ªèTv6iOi)
    sql = sql & ", SGCK6OIL"                '' ÐckTv6iOiá)
    sql = sql & ", SGCK6OIM"                '' ÐckTv6iOi)
    sql = sql & ", SGCK6OIH"                '' ÐckTv6iOi)
    sql = sql & ", CVOIL"                   '' CViOiá)
    sql = sql & ", CVOIM"                   '' CViOi)
    sql = sql & ", CVOIH"                   '' CViOi)
    sql = sql & ", TSTAFFID"                '' o^ÐõID
    sql = sql & ", REGDATE"                 '' o^út
    sql = sql & ", KSTAFFID"                '' XVÐõID
    sql = sql & ", UPDDATE"                 '' XVút
    sql = sql & ", SENDFLAG"                '' MtO
    sql = sql & ", SENDDATE"                '' Mút
    sql = sql & ")"
    
    sql = sql & "Values("
    sql = sql & "'" & record.GOUKI & "'"                                        '' @
    sql = sql & ", " & "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss')"     '' út
    sql = sql & ", " & record.FTIROIL                                           '' FTIRiOiá)
    sql = sql & ", " & record.FTIROIM                                           '' FTIRiOi)
    sql = sql & ", " & record.FTIROIH                                           '' FTIRiOi)
    sql = sql & ", " & record.MS1OIL                                            '' ªèTv1iOiá)
    sql = sql & ", " & record.MS1OIM                                            '' ªèTv1iOi)
    sql = sql & ", " & record.MS1OIH                                            '' ªèTv1iOi)
    sql = sql & ", " & record.MS2OIL                                            '' ªèTv2iOiá)
    sql = sql & ", " & record.MS2OIM                                            '' ªèTv2iOi)
    sql = sql & ", " & record.MS2OIH                                            '' ªèTv2iOi)
    sql = sql & ", " & record.MS3OIL                                            '' ªèTv3iOiá)
    sql = sql & ", " & record.MS3OIM                                            '' ªèTv3iOi)
    sql = sql & ", " & record.MS3OIH                                            '' ªèTv3iOi)
    sql = sql & ", " & record.MS4OIL                                            '' ªèTv4iOiá)
    sql = sql & ", " & record.MS4OIM                                            '' ªèTv4iOi)
    sql = sql & ", " & record.MS4OIH                                            '' ªèTv4iOi)
    sql = sql & ", " & record.MS5OIL                                            '' ªèTv5iOiá)
    sql = sql & ", " & record.MS5OIM                                            '' ªèTv5iOi)
    sql = sql & ", " & record.MS5OIH                                            '' ªèTv5iOi)
    sql = sql & ", " & record.MSAVEOIL                                          '' ªè½ÏiOiá)
    sql = sql & ", " & record.MSAVEOIM                                          '' ªè½ÏiOi)
    sql = sql & ", " & record.MSAVEOIH                                          '' ªè½ÏiOi)
    sql = sql & ", " & record.MSSGOIL                                           '' ªèÐiOiá)
    sql = sql & ", " & record.MSSGOIM                                           '' ªèÐiOi)
    sql = sql & ", " & record.MSSGOIH                                           '' ªèÐiOi)
    sql = sql & ", " & record.MSPSGOIL                                          '' ªèAVE+ÐiOiá)
    sql = sql & ", " & record.MSPSGOIM                                          '' ªèAVE+ÐiOi)
    sql = sql & ", " & record.MSPSGOIH                                          '' ªèAVE+ÐiOi)
    sql = sql & ", " & record.MSNSGOIL                                          '' ªèAVE-ÐiOiá)
    sql = sql & ", " & record.MSNSGOIM                                          '' ªèAVE-ÐiOi)
    sql = sql & ", " & record.MSNSGOIH                                          '' ªèAVE-ÐiOi)
    sql = sql & ", " & record.MINOIL                                            '' MINiOiá)
    sql = sql & ", " & record.MINOIM                                            '' MINiOi)
    sql = sql & ", " & record.MINOIH                                            '' MINiOi)
    sql = sql & ", " & record.MAXOIL                                            '' MAXiOiá)
    sql = sql & ", " & record.MAXOIM                                            '' MAXiOi)
    sql = sql & ", " & record.MAXOIH                                            '' MAXiOi)
    sql = sql & ", " & record.SGCK1OIL                                          '' ÐckTv1iOiá)
    sql = sql & ", " & record.SGCK1OIM                                          '' ÐckTv1iOi)
    sql = sql & ", " & record.SGCK1OIH                                          '' ÐckTv1iOi)
    sql = sql & ", " & record.SGCK2OIL                                          '' ÐckTv2iOiá)
    sql = sql & ", " & record.SGCK2OIM                                          '' ÐckTv2iOi)
    sql = sql & ", " & record.SGCK2OIH                                          '' ÐckTv2iOi)
    sql = sql & ", " & record.SGCK3OIL                                          '' ÐckTv3iOiá)
    sql = sql & ", " & record.SGCK3OIM                                          '' ÐckTv3iOi)
    sql = sql & ", " & record.SGCK3OIH                                          '' ÐckTv3iOi)
    sql = sql & ", " & record.SGCK4OIL                                          '' ÐckTv4iOiá)
    sql = sql & ", " & record.SGCK4OIM                                          '' ÐckTv4iOi)
    sql = sql & ", " & record.SGCK4OIH                                          '' ÐckTv4iOi)
    sql = sql & ", " & record.SGCK5OIL                                          '' ÐckTv5iOiá)
    sql = sql & ", " & record.SGCK5OIM                                          '' ÐckTv5iOi)
    sql = sql & ", " & record.SGCK5OIH                                          '' ÐckTv5iOi)
    sql = sql & ", " & record.SGCKDOIL                                          '' Ðckf[^iOiá)
    sql = sql & ", " & record.SGCKDOIM                                          '' Ðckf[^iOi)
    sql = sql & ", " & record.SGCKDOIH                                          '' Ðckf[^iOi)
    sql = sql & ", " & record.SGCKAOIL                                          '' Ðck½ÏiOiá)
    sql = sql & ", " & record.SGCKAAOIM                                         '' Ðck½ÏiOi)
    sql = sql & ", " & record.SGCKAOIH                                          '' Ðck½ÏiOi)
    sql = sql & ", " & record.SGNOIL                                            '' ÐckÐiOiá)
    sql = sql & ", " & record.SGNOIM                                            '' ÐckÐiOi)
    sql = sql & ", " & record.SGNOIH                                            '' ÐckÐiOi)
    sql = sql & ", " & record.FTIRKOIL                                          '' FTIR·ZiOiá)
    sql = sql & ", " & record.FTIRKOIM                                          '' FTIR·ZiOi)
    sql = sql & ", " & record.FTIRKOIH                                          '' FTIR·ZiOi)
    sql = sql & ", " & record.EFFECTTM                                          '' LøÔ
    sql = sql & ", " & record.YCOEF                                             '' eshq·Z®ixØÐj
    sql = sql & ", " & record.XCOEF                                             '' eshq·Z®iwWj
    sql = sql & ", " & record.RSQUARE                                           '' qQæ
    sql = sql & ", " & record.SGCKST                                            '' Ð»èî
    sql = sql & ", '" & record.SGCKOIL & "'"                                    '' Ð»èiOiá)
    sql = sql & ", '" & record.SGCKOIM & "'"                                    '' Ð»èiOi)
    sql = sql & ", '" & record.SGCKOIH & "'"                                    '' Ð»èiOi)
    sql = sql & ", " & record.FTIRCKST                                          '' FTIR·Z»èî
    sql = sql & ", '" & record.FTIRCKOIL & "'"                                  '' FTIR·Z»èiOiá)
    sql = sql & ", '" & record.FTIRCKOIM & "'"                                  '' FTIR·Z»èiOi)
    sql = sql & ", '" & record.FTIRCKOIH & "'"                                  '' FTIR·Z»èiOi)
    sql = sql & ", " & record.MS6OIL                                            '' ªèTv6iOiá)
    sql = sql & ", " & record.MS6OIM                                            '' ªèTv6iOi)
    sql = sql & ", " & record.MS6OIH                                            '' ªèTv6iOi)
    sql = sql & ", " & record.SGCK6OIL                                          '' ÐckTv6iOiá)
    sql = sql & ", " & record.SGCK6OIM                                          '' ÐckTv6iOi)
    sql = sql & ", " & record.SGCK6OIH                                          '' ÐckTv6iOi)
    sql = sql & ", " & record.CVOIL                                             '' CViOiá)
    sql = sql & ", " & record.CVOIM                                             '' CViOi)
    sql = sql & ", " & record.CVOIH                                             '' CViOi)
    sql = sql & ", '" & TSTAFFID & "'"                                          '' o^ÐõID
    sql = sql & ", SYSDATE"                                                     '' o^út
    sql = sql & ", ' '"                                                         '' XVÐõID
    sql = sql & ", SYSDATE"                                                     '' XVút
    sql = sql & ", '0'"                                                         '' MtO
    sql = sql & ", SYSDATE"                                                     '' Mút
    sql = sql & ")"
  
    '' ¡SQLÌÀs
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001j_Exec = FUNCTION_RETURN_SUCCESS

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

'Tv      :f[^Ï·ðs¤
'Êß×Ò°À    :Ï¼        ,IO ,^                ,à¾
'          :tblLeft       ,IO   ,typ_TBCMB019      ,e[uf[^P
'          :tblRight      ,IO   ,typ_cmjc001j_Disp ,e[uf[^Q
'          :bFlg          ,I   ,Boolean           ,TRUE:øPf[^¨øQf[^ÖÌÏ·  FALSE:øPf[^©øQf[^ÖÌÏ·
'à¾      :
Public Sub ConvDate_F_cmjc001j_a(tblLeft As typ_TBCMB019, tblRight As typ_cmjc001j_Disp, bFlg As Boolean)
    
    If bFlg = True Then
        With tblRight
            .GOUKI = tblLeft.GOUKI
            .INPDATE = tblLeft.INPDATE
            .FTIROIL = tblLeft.FTIROIL
            .FTIROIM = tblLeft.FTIROIM
            .FTIROIH = tblLeft.FTIROIH
            .MS1OIL = tblLeft.MS1OIL
            .MS1OIM = tblLeft.MS1OIM
            .MS1OIH = tblLeft.MS1OIH
            .MS2OIL = tblLeft.MS2OIL
            .MS2OIM = tblLeft.MS2OIM
            .MS2OIH = tblLeft.MS2OIH
            .MS3OIL = tblLeft.MS3OIL
            .MS3OIM = tblLeft.MS3OIM
            .MS3OIH = tblLeft.MS3OIH
            .MS4OIL = tblLeft.MS4OIL
            .MS4OIM = tblLeft.MS4OIM
            .MS4OIH = tblLeft.MS4OIH
            .MS5OIL = tblLeft.MS5OIL
            .MS5OIM = tblLeft.MS5OIM
            .MS5OIH = tblLeft.MS5OIH
            .MSAVEOIL = tblLeft.MSAVEOIL
            .MSAVEOIM = tblLeft.MSAVEOIM
            .MSAVEOIH = tblLeft.MSAVEOIH
            .MSSGOIL = tblLeft.MSSGOIL
            .MSSGOIM = tblLeft.MSSGOIM
            .MSSGOIH = tblLeft.MSSGOIH
            .MSPSGOIL = tblLeft.MSPSGOIL
            .MSPSGOIM = tblLeft.MSPSGOIM
            .MSPSGOIH = tblLeft.MSPSGOIH
            .MSNSGOIL = tblLeft.MSNSGOIL
            .MSNSGOIM = tblLeft.MSNSGOIM
            .MSNSGOIH = tblLeft.MSNSGOIH
            .MINOIL = tblLeft.MINOIL
            .MINOIM = tblLeft.MINOIM
            .MINOIH = tblLeft.MINOIH
            .MAXOIL = tblLeft.MAXOIL
            .MAXOIM = tblLeft.MAXOIM
            .MAXOIH = tblLeft.MAXOIH
            .SGCK1OIL = tblLeft.SGCK1OIL
            .SGCK1OIM = tblLeft.SGCK1OIM
            .SGCK1OIH = tblLeft.SGCK1OIH
            .SGCK2OIL = tblLeft.SGCK2OIL
            .SGCK2OIM = tblLeft.SGCK2OIM
            .SGCK2OIH = tblLeft.SGCK2OIH
            .SGCK3OIL = tblLeft.SGCK3OIL
            .SGCK3OIM = tblLeft.SGCK3OIM
            .SGCK3OIH = tblLeft.SGCK3OIH
            .SGCK4OIL = tblLeft.SGCK4OIL
            .SGCK4OIM = tblLeft.SGCK4OIM
            .SGCK4OIH = tblLeft.SGCK4OIH
            .SGCK5OIL = tblLeft.SGCK5OIL
            .SGCK5OIM = tblLeft.SGCK5OIM
            .SGCK5OIH = tblLeft.SGCK5OIH
            .SGCKDOIL = tblLeft.SGCKDOIL
            .SGCKDOIM = tblLeft.SGCKDOIM
            .SGCKDOIH = tblLeft.SGCKDOIH
            .SGCKAOIL = tblLeft.SGCKAOIL
            .SGCKAAOIM = tblLeft.SGCKAAOIM
            .SGCKAOIH = tblLeft.SGCKAOIH
            .SGNOIL = tblLeft.SGNOIL
            .SGNOIM = tblLeft.SGNOIM
            .SGNOIH = tblLeft.SGNOIH
            .FTIRKOIL = tblLeft.FTIRKOIL
            .FTIRKOIM = tblLeft.FTIRKOIM
            .FTIRKOIH = tblLeft.FTIRKOIH
            .EFFECTTM = tblLeft.EFFECTTM
            .YCOEF = tblLeft.YCOEF
            .XCOEF = tblLeft.XCOEF
            .RSQUARE = tblLeft.RSQUARE
            .SGCKST = tblLeft.SGCKST
            .SGCKOIL = tblLeft.SGCKOIL
            .SGCKOIM = tblLeft.SGCKOIM
            .SGCKOIH = tblLeft.SGCKOIH
            .FTIRCKST = tblLeft.FTIRCKST
            .FTIRCKOIL = tblLeft.FTIRCKOIL
            .FTIRCKOIM = tblLeft.FTIRCKOIM
            .FTIRCKOIH = tblLeft.FTIRCKOIH
            .MS6OIL = tblLeft.MS6OIL
            .MS6OIM = tblLeft.MS6OIM
            .MS6OIH = tblLeft.MS6OIH
            .SGCK6OIL = tblLeft.SGCK6OIL
            .SGCK6OIM = tblLeft.SGCK6OIM
            .SGCK6OIH = tblLeft.SGCK6OIH
            .CVOIL = tblLeft.CVOIL
            .CVOIM = tblLeft.CVOIM
            .CVOIH = tblLeft.CVOIH
        
        End With
    Else
        With tblLeft
            .GOUKI = tblRight.GOUKI
            .INPDATE = tblRight.INPDATE
            .FTIROIL = tblRight.FTIROIL
            .FTIROIM = tblRight.FTIROIM
            .FTIROIH = tblRight.FTIROIH
            .MS1OIL = tblRight.MS1OIL
            .MS1OIM = tblRight.MS1OIM
            .MS1OIH = tblRight.MS1OIH
            .MS2OIL = tblRight.MS2OIL
            .MS2OIM = tblRight.MS2OIM
            .MS2OIH = tblRight.MS2OIH
            .MS3OIL = tblRight.MS3OIL
            .MS3OIM = tblRight.MS3OIM
            .MS3OIH = tblRight.MS3OIH
            .MS4OIL = tblRight.MS4OIL
            .MS4OIM = tblRight.MS4OIM
            .MS4OIH = tblRight.MS4OIH
            .MS5OIL = tblRight.MS5OIL
            .MS5OIM = tblRight.MS5OIM
            .MS5OIH = tblRight.MS5OIH
            .MSAVEOIL = tblRight.MSAVEOIL
            .MSAVEOIM = tblRight.MSAVEOIM
            .MSAVEOIH = tblRight.MSAVEOIH
            .MSSGOIL = tblRight.MSSGOIL
            .MSSGOIM = tblRight.MSSGOIM
            .MSSGOIH = tblRight.MSSGOIH
            .MSPSGOIL = tblRight.MSPSGOIL
            .MSPSGOIM = tblRight.MSPSGOIM
            .MSPSGOIH = tblRight.MSPSGOIH
            .MSNSGOIL = tblRight.MSNSGOIL
            .MSNSGOIM = tblRight.MSNSGOIM
            .MSNSGOIH = tblRight.MSNSGOIH
            .MINOIL = tblRight.MINOIL
            .MINOIM = tblRight.MINOIM
            .MINOIH = tblRight.MINOIH
            .MAXOIL = tblRight.MAXOIL
            .MAXOIM = tblRight.MAXOIM
            .MAXOIH = tblRight.MAXOIH
            .SGCK1OIL = tblRight.SGCK1OIL
            .SGCK1OIM = tblRight.SGCK1OIM
            .SGCK1OIH = tblRight.SGCK1OIH
            .SGCK2OIL = tblRight.SGCK2OIL
            .SGCK2OIM = tblRight.SGCK2OIM
            .SGCK2OIH = tblRight.SGCK2OIH
            .SGCK3OIL = tblRight.SGCK3OIL
            .SGCK3OIM = tblRight.SGCK3OIM
            .SGCK3OIH = tblRight.SGCK3OIH
            .SGCK4OIL = tblRight.SGCK4OIL
            .SGCK4OIM = tblRight.SGCK4OIM
            .SGCK4OIH = tblRight.SGCK4OIH
            .SGCK5OIL = tblRight.SGCK5OIL
            .SGCK5OIM = tblRight.SGCK5OIM
            .SGCK5OIH = tblRight.SGCK5OIH
            .SGCKDOIL = tblRight.SGCKDOIL
            .SGCKDOIM = tblRight.SGCKDOIM
            .SGCKDOIH = tblRight.SGCKDOIH
            .SGCKAOIL = tblRight.SGCKAOIL
            .SGCKAAOIM = tblRight.SGCKAAOIM
            .SGCKAOIH = tblRight.SGCKAOIH
            .SGNOIL = tblRight.SGNOIL
            .SGNOIM = tblRight.SGNOIM
            .SGNOIH = tblRight.SGNOIH
            .FTIRKOIL = tblRight.FTIRKOIL
            .FTIRKOIM = tblRight.FTIRKOIM
            .FTIRKOIH = tblRight.FTIRKOIH
            .EFFECTTM = tblRight.EFFECTTM
            .YCOEF = tblRight.YCOEF
            .XCOEF = tblRight.XCOEF
            .RSQUARE = tblRight.RSQUARE
            .SGCKST = tblRight.SGCKST
            .SGCKOIL = tblRight.SGCKOIL
            .SGCKOIM = tblRight.SGCKOIM
            .SGCKOIH = tblRight.SGCKOIH
            .FTIRCKST = tblRight.FTIRCKST
            .FTIRCKOIL = tblRight.FTIRCKOIL
            .FTIRCKOIM = tblRight.FTIRCKOIM
            .FTIRCKOIH = tblRight.FTIRCKOIH
            .MS6OIL = tblRight.MS6OIL
            .MS6OIM = tblRight.MS6OIM
            .MS6OIH = tblRight.MS6OIH
            .SGCK6OIL = tblRight.SGCK6OIL
            .SGCK6OIM = tblRight.SGCK6OIM
            .SGCK6OIH = tblRight.SGCK6OIH
            .CVOIL = tblRight.CVOIL
            .CVOIM = tblRight.CVOIM
            .CVOIH = tblRight.CVOIH
        
        End With
    End If

End Sub

'''''------------------------------------------------
''''' DBANZXÖ
'''''------------------------------------------------
''''
'''''Tv      :e[uuTBCMB019v©çðÉ Á½R[hðo·é
'''''Êß×Ò°À    :Ï¼        ,IO ,^           ,à¾
'''''          :records()     ,O  ,typ_TBCMB019 ,oR[h
'''''          :sqlWhere      ,I  ,String       ,oð(SQLÌWhereß:ÈªÂ\)
'''''          :sqlOrder      ,I  ,String       ,o(SQLÌOrder byß:ÈªÂ\)
'''''          :ßèl        ,O  ,FUNCTION_RETURN ,oÌ¬Û
'''''à¾      :
'''''ð      :2001/08/24ì¬@ìº
''''Public Function DBDRV_GetTBCMB019(records() As typ_TBCMB019, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
''''Dim sql As String       'SQLSÌ
''''Dim sqlBase As String   'SQLî{(WHEREßÌOÜÅ)
''''Dim rs As OraDynaset    'RecordSet
''''Dim recCnt As Long      'R[h
''''Dim i As Long
''''
''''    ''SQLðgÝ§Äé
''''    sqlBase = "Select GOUKI, INPDATE, FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
''''              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
''''              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
''''              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
''''              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
''''              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG," & _
''''              " SENDDATE "
''''    sqlBase = sqlBase & "From TBCMB019"
''''    sql = sqlBase
''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
''''        sql = sql & " " & sqlWhere & " " & sqlOrder
''''    End If
''''
''''    ''f[^ðo·é
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''    If rs Is Nothing Then
''''        ReDim records(0)
''''        DBDRV_GetTBCMB019 = FUNCTION_RETURN_FAILURE
''''        Exit Function
''''    End If
''''
''''    ''oÊði[·é
''''    recCnt = rs.RecordCount
''''    ReDim records(recCnt)
''''    For i = 1 To recCnt
''''        With records(i)
''''            .GOUKI = rs("GOUKI")             ' @
''''            .INPDATE = rs("INPDATE")         ' út
''''            .FTIRFZI = rs("FTIRFZI")         ' FTIRiFZ)
''''            .FTIRCZH = rs("FTIRCZH")         ' FTIRiCZj
''''            .FTIRCZC = rs("FTIRCZC")         ' FTIRiCZj
''''            .MS1FZ = rs("MS1FZ")             ' ªèTv1iFZ)
''''            .MS1CZ1 = rs("MS1CZ1")           ' ªèTv1iCZ-1)
''''            .MS1CZ2 = rs("MS1CZ2")           ' ªèTv1iCZ-2)
''''            .MS2FZ = rs("MS2FZ")             ' ªèTv2iFZ)
''''            .MS2CZ1 = rs("MS2CZ1")           ' ªèTv2iCZ-1)
''''            .MS2CZ2 = rs("MS2CZ2")           ' ªèTv2iCZ-2)
''''            .MS3FZ = rs("MS3FZ")             ' ªèTv3iFZ)
''''            .MS3CZ1 = rs("MS3CZ1")           ' ªèTv3iCZ-1)
''''            .MS3CZ2 = rs("MS3CZ2")           ' ªèTv3iCZ-2)
''''            .MS4FZ = rs("MS4FZ")             ' ªèTv4iFZ)
''''            .MS4CZ1 = rs("MS4CZ1")           ' ªèTv4iCZ-1)
''''            .MS4CZ2 = rs("MS4CZ2")           ' ªèTv4iCZ-2)
''''            .MS5FZ = rs("MS5FZ")             ' ªèTv5iFZ)
''''            .MS5CZ1 = rs("MS5CZ1")           ' ªèTv5iCZ-1)
''''            .MS5CZ2 = rs("MS5CZ2")           ' ªèTv5iCZ-2)
''''            .MSAVEFZ = rs("MSAVEFZ")         ' ªè½ÏiFZj
''''            .MSAVECZ1 = rs("MSAVECZ1")       ' ªè½ÏiCZ-1j
''''            .MSAVECZ2 = rs("MSAVECZ2")       ' ªè½ÏiCZ-2j
''''            .MSSGFZ = rs("MSSGFZ")           ' ªèÐiFZj
''''            .MSSGCZ1 = rs("MSSGCZ1")         ' ªèÐiCZ-1j
''''            .MSSGCZ2 = rs("MSSGCZ2")         ' ªèÐiCZ-2j
''''            .MSPSGFZ = rs("MSPSGFZ")         ' ªèAVE+ÐiFZj
''''            .MSPSGCZ1 = rs("MSPSGCZ1")       ' ªèAVE+ÐiCZ-1j
''''            .MSPSGCZ2 = rs("MSPSGCZ2")       ' ªèAVE+ÐiCZ-2j
''''            .MSNSGFZ = rs("MSNSGFZ")         ' ªèAVE-ÐiFZj
''''            .MSNSGCZ1 = rs("MSNSGCZ1")       ' ªèAVE-ÐiCZ-1j
''''            .MSNSGCZ2 = rs("MSNSGCZ2")       ' ªèAVE-ÐiCZ-2j
''''            .MINFZ = rs("MINFZ")             ' MINiFZj
''''            .MINCZ1 = rs("MINCZ1")           ' MINiCZ-1j
''''            .MINCZ2 = rs("MINCZ2")           ' MINiCZ-2j
''''            .MAXFZ = rs("MAXFZ")             ' MAXiFZj
''''            .MAXCZ1 = rs("MAXCZ1")           ' MAXiCZ-1j
''''            .MAXCZ2 = rs("MAXCZ2")           ' MAXiCZ-2j
''''            .SGCK1FZ = rs("SGCK1FZ")         ' ÐckTv1iFZ)
''''            .SGCK1CZ1 = rs("SGCK1CZ1")       ' ÐckTv1iCZ-1)
''''            .SGCK1CZ2 = rs("SGCK1CZ2")       ' ÐckTv1iCZ-2)
''''            .SGCK2FZ = rs("SGCK2FZ")         ' ÐckTv2iFZ)
''''            .SGCK2CZ1 = rs("SGCK2CZ1")       ' ÐckTv2iCZ-1)
''''            .SGCK2CZ2 = rs("SGCK2CZ2")       ' ÐckTv2iCZ-2)
''''            .SGCK3FZ = rs("SGCK3FZ")         ' ÐckTv3iFZ)
''''            .SGCK3CZ1 = rs("SGCK3CZ1")       ' ÐckTv3iCZ-1)
''''            .SGCK3CZ2 = rs("SGCK3CZ2")       ' ÐckTv3iCZ-2)
''''            .SGCK4FZ = rs("SGCK4FZ")         ' ÐckTv4iFZ)
''''            .SGCK4CZ1 = rs("SGCK4CZ1")       ' ÐckTv4iCZ-1)
''''            .SGCK4CZ2 = rs("SGCK4CZ2")       ' ÐckTv4iCZ-2)
''''            .SGCK5FZ = rs("SGCK5FZ")         ' ÐckTv5iFZ)
''''            .SGCK5CZ1 = rs("SGCK5CZ1")       ' ÐckTv5iCZ-1)
''''            .SGCK5CZ2 = rs("SGCK5CZ2")       ' ÐckTv5iCZ-2)
''''            .SGCKDFZ = rs("SGCKDFZ")         ' Ðckf[^iFZj
''''            .SGCKDCZ1 = rs("SGCKDCZ1")       ' Ðckf[^iCZ-1j
''''            .SGCKDCZ2 = rs("SGCKDCZ2")       ' Ðckf[^iCZ-2j
''''            .SGCKAFZ = rs("SGCKAFZ")         ' Ðck½ÏiFZj
''''            .SGCKAACZ1 = rs("SGCKAACZ1")     ' Ðck½ÏiCZ-1j
''''            .SGCKACZ2 = rs("SGCKACZ2")       ' Ðck½ÏiCZ-2j
''''            .SGNFZ = rs("SGNFZ")             ' ÐckÐiFZj
''''            .SGNCZ1 = rs("SGNCZ1")           ' ÐckÐ CZ-1j
''''            .SGNCZ2 = rs("SGNCZ2")           ' ÐckÐiCZ-2j
''''            .FTIRFZ = rs("FTIRFZ")           ' FTIR·ZiFZj
''''            .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR·ZiCZ-1j
''''            .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR·ZiCZ-2j
''''            .EFFECTTM = rs("EFFECTTM")       ' LøÔ
''''            .YCOEF = rs("YCOEF")             ' eshq·Z®ixØÐj
''''            .XCOEF = rs("XCOEF")             ' eshq·Z®iwWj
''''            .RSQUARE = rs("RSQUARE")         ' qQæ
''''            .TSTAFFID = rs("TSTAFFID")       ' o^ÐõID
''''            .REGDATE = rs("REGDATE")         ' o^út
''''            .KSTAFFID = rs("KSTAFFID")       ' XVÐõID
''''            .UPDDATE = rs("UPDDATE")         ' XVút
''''            .SENDFLAG = rs("SENDFLAG")       ' MtO
''''            .SENDDATE = rs("SENDDATE")       ' Mút
''''        End With
''''        rs.MoveNext
''''    Next
''''    rs.Close
''''
''''    DBDRV_GetTBCMB019 = FUNCTION_RETURN_SUCCESS
''''End Function

'///////////////////////////////////////////////////
' @(f)
' @\    : Ð»èîæ¾
'
' Ôèl  : True  - ³í
' @@@    False - ¸s
'
' ø«  : sSigCode  - Ð»èî
' @@@  : sFtirCode - FTIR·Z»èî
' @@@  : sR2Code   - R2æ»èî
'
' @\à¾:
'///////////////////////////////////////////////////
Public Function GetSigChkCode(Optional ByRef sSigCode As String _
                            , Optional ByRef sFtirCode As String _
                            , Optional ByRef sR2Code As String _
                            ) As Boolean
    Dim dbIsMine    As Boolean
    Dim sSql        As String
    Dim objRs       As Object
    
    GetSigChkCode = False
    sSigCode = ""
    sFtirCode = ""
    
    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "s_cmzc055_SQL.bas -- Function GetSigChkCode"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''rpk¶ì¬
    sSql = ""
    sSql = sSql & "SELECT NVL(kcode01a9, ' ')"   '0:Ð»èî
    sSql = sSql & "      ,NVL(kcode02a9, ' ')"   '1:FTIR·Z»èî
    sSql = sSql & "      ,NVL(kcode03a9, ' ')"   '2:R2æ»èî
    sSql = sSql & "  FROM koda9"
    sSql = sSql & " WHERE sysca9 = 'X'"
    sSql = sSql & "   AND shuca9 = '19'"
    sSql = sSql & "   AND codea9 = 'FRS'"
    
    Set objRs = OraDB.CreateDynaset(sSql, ORADYN_DEFAULT)
    
    If objRs.EOF Then
        Call MsgOut(0, "Ð»èîÌR[hªo^³êÄ¢Ü¹ñ", ERR_DISP)
        Exit Function
    End If

    sSigCode = objRs(0)     ''Ð»èî
    sFtirCode = objRs(1)    ''FTIR·Z»èî
    sR2Code = objRs(2)      ''R2æ»èî
    
    objRs.Close
    
    ''Ð»èî
    If IsNumeric(sSigCode) = False Then
        Call MsgOut(0, "Ð»èîÌR[hª³µ­ èÜ¹ñ", ERR_DISP)
        Exit Function
    End If
    ' -10~100ÅÈ¢CÜ½Í¬_æOÊÈ~ÌüÍª éêÍG[
    If Not (-10# < CDbl(sSigCode) And CDbl(sSigCode) < 100#) Then
        Call MsgOut(0, "Ð»èîÌR[hª³µ­ èÜ¹ñ", ERR_DISP)
        Exit Function
    End If
    If InStr(1, sSigCode, ".", vbTextCompare) >= 1 Then
        If Len(sSigCode) - InStr(1, sSigCode, ".", vbTextCompare) >= 3 Then
            Call MsgOut(0, "Ð»èîÌR[hª³µ­ èÜ¹ñ", ERR_DISP)
            Exit Function
        End If
    End If
    
    ''FTIR·Z»èî
    If IsNumeric(sFtirCode) = False Then
        Call MsgOut(0, "FTIR·Z»èîÌR[hª³µ­ èÜ¹ñ", ERR_DISP)
        Exit Function
    End If
    ' -10~100ÅÈ¢CÜ½Í¬_æOÊÈ~ÌüÍª éêÍG[
    If Not (-10# < CDbl(sFtirCode) And CDbl(sFtirCode) < 100#) Then
        Call MsgOut(0, "FTIR·Z»èîÌR[hª³µ­ èÜ¹ñ", ERR_DISP)
        Exit Function
    End If
    If InStr(1, sFtirCode, ".", vbTextCompare) >= 1 Then
        If Len(sFtirCode) - InStr(1, sFtirCode, ".", vbTextCompare) >= 3 Then
            Call MsgOut(0, "FTIR·Z»èîÌR[hª³µ­ èÜ¹ñ", ERR_DISP)
            Exit Function
        End If
    End If
    
    ''R2æ»èî
    If IsNumeric(sR2Code) = False Then
        Call MsgOut(0, "q2æ»èîÌR[hª³µ­ èÜ¹ñ", ERR_DISP)
        Exit Function
    End If
    
    GetSigChkCode = True        ''¬÷ðÔ·

proc_exit:
    If dbIsMine Then
        OraDBClose
    End If
    
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    gErr.HandleError
    Resume proc_exit
    
End Function

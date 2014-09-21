Attribute VB_Name = "s_cmbc007_SQL"
Option Explicit

'´¿ÝÉC³


'´¿ÝÉæ¾p
Public Type type_DBDRV_scmzc_fcmgc001e_Disp
    '´¿ÝÉÇ
    MTRLNUM As String * 10      ' ´¿Ô
    WEIGHT As Long              ' dÊ
End Type


'´¿ÝÉXVp
Public Type type_DBDRV_scmzc_fcmgc001e_Exec
    '´¿ÝÉÇ
    MTRLNUM As String * 10      ' ´¿Ô
    USABLCLS As String * 1      ' gpÂ\æª
    KRPROCCD As String * 5      ' ÇHöR[h
    PROCCODE As String * 5      ' HöR[h
    KSTAFFID As String * 8      ' XVÐõID
    WEIGHT As Long              ' VdÊ
    SYORIW As Long              ' Ê

End Type



'ú\¦
'Tv    :´¿ÝÉC³ \¦pcahCo
'Êß×Ò°À  :Ï¼       ,IO  ,^                                    ,à¾
'        :record       ,I   ,type_DBDRV_scmzc_fcmgc001e_Disp       ,´¿ÝÉæ¾p
'        :ßØl        ,O   ,FUNCTION_RETURN                       ,ÇÝÝ¬Û
'à¾    :
'ð    :2001/06/18  { ì¬
Public Function DBDRV_scmzc_fcmgc001e_Disp(records() As type_DBDRV_scmzc_fcmgc001e_Disp) As FUNCTION_RETURN
    
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Integer
    Dim i As Long
    
    '´¿Çe[uÅgpÂ\æªPðselecti´¿ÔAdÊj
    

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmgc001e_SQL.bas -- Function DBDRV_scmzc_fcmgc001e_Disp"

    sql = "Select MTRLNUM, USABLCLS, WEIGHT, TSTAFFID, REGDATE, KSTAFFID, UPDDATE "
    sql = sql & "From TBCMG005"
    
        
        sql = "select MTRLNUM, WEIGHT"
        sql = sql & " from ( "
        sql = sql & " select MTRLNUM, WEIGHT"
        sql = sql & " from TBCMG005"
        sql = sql & " where USABLCLS='1'"
        sql = sql & " and WEIGHT > 0 "
        sql = sql & " and substr(MTRLNUM,1,1) not in ('P','N')"
        sql = sql & " order by MTRLNUM ) "
        sql = sql & " union all "
        sql = sql & " select MTRLNUM, WEIGHT"
        sql = sql & " from ( "
        sql = sql & " select MTRLNUM, WEIGHT"
        sql = sql & " from TBCMG005"
        sql = sql & " where USABLCLS='1'"
        sql = sql & " and WEIGHT > 0 "
        sql = sql & " and substr(MTRLNUM,1,1) in ('P','N')"
        sql = sql & " order by MTRLNUM )"


    '   order by ´¿Ô
    '   substr(´¿Ô,1,1) not in ('P','N')
    '   union all
    
    'select ...
    '   order by ´¿Ô
    '   substr(´¿Ô,1,1) in ('P','N')
    
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    recCnt = rs.RecordCount
    ReDim records(recCnt)
    If recCnt = 0 Then ''2001/07/17 Sano
'2001/07/17 Sano    If rs.RecordCount = 0 Then
        DBDRV_scmzc_fcmgc001e_Disp = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If
    
'2001/07/17 Sano    recCnt = rs.RecordCount
'2001/07/17 Sano    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .MTRLNUM = rs("MTRLNUM")          ' ´¿Ô
            .WEIGHT = rs("WEIGHT")            ' dÊ
        End With
        rs.MoveNext
    Next i
    rs.Close

    DBDRV_scmzc_fcmgc001e_Disp = FUNCTION_RETURN_SUCCESS
   

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmgc001e_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'Às
'Tv    :´¿ÝÉC³ XVA}üpcahCo
'Êß×Ò°À  :Ï¼       ,IO  ,^                                    ,à¾
'        :record       ,I   ,type_DBDRV_scmzc_fcmgc001e_Exec       ,´¿ÝÉ}üp
'        :ßØl        ,O   ,FUNCTION_RETURN                       ,ÇÝÝ¬Û
'à¾    :
'ð    :2001/06/18  { ì¬
Public Function DBDRV_scmzc_fcmgc001e_Exec(record As type_DBDRV_scmzc_fcmgc001e_Exec) As FUNCTION_RETURN
    
    Dim sql As String

    
    '´¿Çe[uðVdÊÉXV
        

    'G[nhÌÝè
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmgc001e_SQL.bas -- Function DBDRV_scmzc_fcmgc001e_Exec"

    DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_SUCCESS
    
    sql = "update TBCMG005 set "
    With record
        sql = sql & "WEIGHT=" & .WEIGHT & ", "               ' dÊ
        sql = sql & "KSTAFFID='" & .KSTAFFID & "', "         ' XVÐõID
        sql = sql & "UPDDATE=sysdate "                       ' XVút
        sql = sql & "where MTRLNUM='" & .MTRLNUM & "' "
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    
    '´¿ÝÉÀÑÉ}ü
    sql = " insert into TBCMG006 ( "
    sql = sql & "MTRLNUM, "          ' ´¿Ô
    sql = sql & "TRANCNT, "          ' ñ
    sql = sql & "KRPROCCD, "         ' ÇHöR[h
    sql = sql & "PROCCODE, "         ' HöR[h
    sql = sql & "CLASS, "            ' æª
    sql = sql & "INWEIGHT, "         ' üÍdÊ
    sql = sql & "TSTAFFID, "         ' o^ÐõID
    sql = sql & "REGDATE, "          ' o^út
    sql = sql & "KSTAFFID, "         ' XVÐõID
    sql = sql & "UPDDATE, "          ' XVút
    sql = sql & "SENDFLAG, "         ' MtO
    sql = sql & "SENDDATE ) "        ' Mút
    With record
        sql = sql & " select "
        sql = sql & " '" & .MTRLNUM & "', "          ' ´¿Ô
        sql = sql & " nvl(max(TRANCNT),0)+1, "       ' ñ
        sql = sql & " '" & .KRPROCCD & "', "         ' ÇHöR[h
        sql = sql & " '" & .PROCCODE & "', "         ' HöR[h
        sql = sql & " '" & .USABLCLS & "', "         ' æª
        sql = sql & " '" & .SYORIW & "', "           ' üÍdÊ
        sql = sql & " '" & .KSTAFFID & "', "         ' o^ÐõID
        sql = sql & " sysdate, "                     ' o^út
        sql = sql & " '" & .KSTAFFID & "', "         ' XVÐõID
        sql = sql & " sysdate, "                     ' XVút
        sql = sql & " '0', "                         ' MtO
        sql = sql & " sysdate "                      ' Mút
        sql = sql & " from TBCMG006 "
        sql = sql & " where MTRLNUM='" & .MTRLNUM & "' "
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
        

proc_exit:
    'I¹
    gErr.Pop
    Exit Function

proc_err:
    'G[nh
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


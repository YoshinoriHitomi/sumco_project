Attribute VB_Name = "f_cmgc001b_SQL"
Option Explicit
'
'' ½»óüÀÑ
'Public Type typ_TBCMG001
'    MTRLNUM As String * 10      ' ´¿Ô
'    JDATE As Date               ' út
'    TRANCNT As Integer          ' ñ
'    KRPROCCD As String * 5      ' ÇHöR[h
'    PROCCODE As String * 6      ' HöR[h
'    MTRLTYPE As String * 3      ' ´¿íÞ
'    MAKERNO As String * 6       ' [JÇNo
'    RVWEIGHT As Double          ' óüwüdÊ
'    CRYCOMMENT As String        ' Rg
'    TSTAFFID As String * 8      ' o^ÐõID
'    REGDATE As Date             ' o^út
'    KSTAFFID As String * 8      ' XVÐõhc
'    UPDDATE As Date             ' XVút
'    SENDFLAG As String * 1      ' MtO
'    SENDDATE As Date            ' Mút
'End Type
'
'' ´¿ÝÉÇ
'Public Type typ_TBCMG005
'    MTRLNUM As String * 10      ' ´¿Ô
'    USABLCLS As String * 1      ' gpÂ\æª
'    WEIGHT As Integer           ' dÊ
'    TSTAFFID As String * 8      ' o^ÐõID
'    REGDATE As Date             ' o^út
'    KSTAFFID As String * 8      ' XVÐõID
'    UPDDATE As Date             ' XVút
'End Type

' f_cmgc001b_Exec
Public Type type_DBDRV_f_cmgc001b_Exec
    KRPROCCD As String * 5      ' ÇHöR[h
    PROCCODE As String * 6      ' HöR[h
    TSTAFFID As String * 8      ' o^ÐõID
    MTRLTYPE As String * 3      ' ´¿íÞ
    MAKERNO As String * 6       ' [JÇNo
    RVWEIGHT As Double          ' óüwüdÊ
    CRYCOMMENT As String        ' Rg
End Type

Public Function DBDRV_f_cmgc001b_Exec(DBDRV_f_cmgc001b_Exec As type_DBDRV_f_cmgc001b_Exec) As FUNCTION_RETURN

    f_cmgc001b_Exec = FUNCTION_RETURN_SUCCESS
End Function

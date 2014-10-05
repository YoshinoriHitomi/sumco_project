Attribute VB_Name = "GetHeatmentNo"
'äTóv    :îMèàóùî‘çÜ
' ﬂ◊“∞¿  :ïœêîñº       ,IO  ,å^                                    ,ê‡ñæ
'        :rcvHinban       ,I   ,tFullHinban         ,ì¸óÕóp
'        : rcvHyouka     ,I   ,String         ,ì¸óÕóp
'        :rcvNetsu        ,I   ,String         ,ì¸óÕóp
'ê‡ñæ    :
'óöó    :2001/06/29 è¨ó— çÏê¨

Public Function GetHeatmentNo(rcvHinban As tFullHinban, rcvHyouka As String, _
                         rcvNetsu As String) As Integer

        Dim sql As String
        Dim NETSU(4) As String
       Dim ii As Integer
        Dim num As Integer

        GetHeatmentNo = 0

        Select Case rcvHyouka
        Case "B"
                sql = "select " & _
                        "HWFBM1NS, HWFBM2NS, HWFBM3NS" & _
                        " from TBCME029" & _
                        " where HINBAN = '" & rcvHinban.hinban & "' and" & _
                        " REVNUM = " & rcvHinban.REVNUM & " and" & _
                        " FACTORY = '" & rcvHinban.FACTORY & "' and" & _
                        " OPECOND = '" & rcvHinban.OPECOND & "'"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                        rs.Close
                        Exit Function
                End If
                NETSU(1) = rs("HWFBM1NS")
                NETSU(2) = rs("HWFBM2NS")
                NETSU(3) = rs("HWFBM3NS")
                num = 3
        Case "DO"
                sql = "select " & _
                        "HWFOS1NS, HWFOS2NS, HWFOS3NS" & _
                        " from TBCME025" & _
                        " where HINBAN = '" & rcvHinban.hinban & "' and" & _
                        " REVNUM = " & rcvHinban.REVNUM & " and" & _
                        " FACTORY = '" & rcvHinban.FACTORY & "' and" & _
                        " OPECOND = '" & rcvHinban.OPECOND & "'"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                        rs.Close
                        Exit Function
                End If
                NETSU(1) = rs("HWFOS1NS")
                NETSU(2) = rs("HWFOS2NS")
                NETSU(3) = rs("HWFOS3NS")
                num = 3
        Case "L"
                sql = "select " & _
                        "HWFOF1NS, HWFOF2NS, HWFOF3NS" & _
                        " from TBCME029" & _
                        " where HINBAN = '" & rcvHinban.hinban & "' and" & _
                        " REVNUM = " & rcvHinban.REVNUM & " and" & _
                        " FACTORY = '" & rcvHinban.FACTORY & "' and" & _
                        " OPECOND = '" & rcvHinban.OPECOND & "'"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                        rs.Close
                        Exit Function
                End If
                NETSU(1) = rs("HWFOF1NS")
                NETSU(2) = rs("HWFOF2NS")
                NETSU(3) = rs("HWFOF3NS")
                NETSU(4) = rs("HWFOF4NS")
                num = 4
        Case Else
                Exit Function
        End Select

        For ii = 1 To num
                If NETSU(ii) = rcvNetsu Then
                        GetHeatmentNo = ii
                        Exit For
                End If
        Next ii


End Function

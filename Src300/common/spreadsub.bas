Attribute VB_Name = "spreadsub"
'P0********************************************************************************
'P0* 関数名　 : JmSortList
'P0* 処理概要 : 指定された取引区分により最新ＮＯを反映する。
'P0* 引数　　 : plngCol As Long                         : 並び替える列
'P0* 戻り値   : 無し
'P0********************************************************************************
'Sub sprSort(ByRef pobjSpread As vaSpread, ByVal plngCol As Long)
Sub sprSort(ByRef pobjSpread As Object, ByVal plngCol As Long)
    On Error GoTo Err

    ' --< 指定された列をキーにソート >--'
'    With sprTest1
    With pobjSpread
        .BlockMode = True                               '  セルブロックを有効
        .Col = 1                                        '  列を設定
        .Col2 = .MaxCols                                '  最終列を設定
        .Row = 1                                        '  行を設定
        .Row2 = .MaxRows                                '  最終行を設定
        .SortBy = SortByRow                             '  行単位に並び替え
        .SortKey(1) = plngCol                           '  並び替えのキーを設定
        
        If .SortKey(1) = plngCol And .SortKeyOrder(1) = SortKeyOrderAscending Then
            .SortKeyOrder(1) = SortKeyOrderDescending   '  降順に並び替えを設定
        Else
            .SortKeyOrder(1) = SortKeyOrderAscending    '  昇順に並び替えを設定
        End If
        
        .Action = ActionSort                            '  並び替えを実行
        .BlockMode = False                              '  セルブロックを無効
    End With

    Exit Sub
    
Err:
    ' --< エラー処理 >--'
'    Call clsCommon.JmRaiseErr(Err, "(clsSpread / JmSortList)")
MsgBox "sprsort err"
End Sub


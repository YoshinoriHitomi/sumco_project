Attribute VB_Name = "spreadsub"
'P0********************************************************************************
'P0* �֐����@ : JmSortList
'P0* �����T�v : �w�肳�ꂽ����敪�ɂ��ŐV�m�n�𔽉f����B
'P0* �����@�@ : plngCol As Long                         : ���ёւ����
'P0* �߂�l   : ����
'P0********************************************************************************
'Sub sprSort(ByRef pobjSpread As vaSpread, ByVal plngCol As Long)
Sub sprSort(ByRef pobjSpread As Object, ByVal plngCol As Long)
    On Error GoTo Err

    ' --< �w�肳�ꂽ����L�[�Ƀ\�[�g >--'
'    With sprTest1
    With pobjSpread
        .BlockMode = True                               '  �Z���u���b�N��L��
        .Col = 1                                        '  ���ݒ�
        .Col2 = .MaxCols                                '  �ŏI���ݒ�
        .Row = 1                                        '  �s��ݒ�
        .Row2 = .MaxRows                                '  �ŏI�s��ݒ�
        .SortBy = SortByRow                             '  �s�P�ʂɕ��ёւ�
        .SortKey(1) = plngCol                           '  ���ёւ��̃L�[��ݒ�
        
        If .SortKey(1) = plngCol And .SortKeyOrder(1) = SortKeyOrderAscending Then
            .SortKeyOrder(1) = SortKeyOrderDescending   '  �~���ɕ��ёւ���ݒ�
        Else
            .SortKeyOrder(1) = SortKeyOrderAscending    '  �����ɕ��ёւ���ݒ�
        End If
        
        .Action = ActionSort                            '  ���ёւ������s
        .BlockMode = False                              '  �Z���u���b�N�𖳌�
    End With

    Exit Sub
    
Err:
    ' --< �G���[���� >--'
'    Call clsCommon.JmRaiseErr(Err, "(clsSpread / JmSortList)")
MsgBox "sprsort err"
End Sub


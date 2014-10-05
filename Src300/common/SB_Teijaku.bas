Attribute VB_Name = "SB_Teijaku"
'-----------------------------------------------------------------------------------------------------------------
'       ��ڃJ�b�g�Ή����ʊ֐����W���[��
'
'
'                                                   �쐬��      20003.09.10
'                                                   �ύX��      20003.XX.XX
'
'
'
'   �쐬��         �V�X�e���u���C��
'
'
'
'-----------------------------------------------------------------------------------------------------------------

Option Explicit
Option Base 1

Public Enum EnumCutFlag
    DoCut = 1
    NoCut = 0
End Enum


Public Type typ_CutBlkHinban
    INPOS As Integer
    Cut As EnumCutFlag
    hinban As tFullHinban
    LENGTH As Integer
End Type


Public Type typ_ChangeHinb
    INPOS As Integer
    Cut As EnumCutFlag
    hinban As tFullHinban
    LENGTH As Integer
End Type


'====================================================================================================================
' �E�w�肳�ꂽ�����̏�񂩂�(�ؒf��)�A��ڃJ�b�g���s�Ȃ����ۂ̃u���b�N�ؒf�����Ăяo�����ɕԂ��B                        *
' �E�i�Ԃ̎d�l���񓙂̗��R�ɂ��A��ڃJ�b�g���s�\�ȏꍇ�A��ڃJ�b�g�s��Ԃ��B                                         *
' �E���i�̐ؒf�̈�敪�ɂ��A�u�P�ƕ��ʂƂ��Đؒf�v�̏ꍇ�A��ڃJ�b�g�\�ȘA���̈�̂ݑΏۂƂ��A                          *
' �E�P�ƕ��ʂ͕ύX�Ȃ��Ƃ���B                                                                                         *
' �E�ŉ��ʕ��̒�����100mm�����̏ꍇ�A100mm�ȏ�ɂȂ�悤�ɒ������s�Ȃ��B                                                 *
'====================================================================================================================
'  �Q�ƃe�[�u��                         TBCME036                                                                    *
'  ���ږ��@�@�@��ۯ��P�ʕۏ��׸�        BLOCKHFLAG  0: �w�肳�ꂽ�A�_���i�Ԃ̎d�l�l���擾����B                         *
'                                                   1: �w�肳��Ă���S�i�Ԃ̎d�l�l���擾����B�i�z��쐬�j             *
'                                                                                                                   *
'   �߂�l                                                                                                          *
'    ����I��               0       �i�Ȃ��j                                                                        *
'    ����I��                                                                                                      *
'   (��ڕs��)              TJ001 ��ۯ���ĕi�Ԃ�����ׁA��ڶ�Ăł��܂���B                                         *
'   �@                      TJ002 �ŉ��ʕ��������擾�ł��Ȃ��פ��ڃJ�b�g�ł��܂���                                *
'                           TJ003 ���グ�� < �ŉ��ʕ������̈פ��ڃJ�b�g�ł��܂���                                 *
'                           TJ004 Z / G�i�Ԃ�TOP / BOT�ȊO�̈פ��ڃJ�b�g�ł��܂���                                *
'                           TJ005 ���ŉ��ʕ��`�F�b�N���s���B�ŉ����ȉ��̃J�b�g��������ꍇ�̓G���[                      *
'                                                                                                                   *
'-------------------------------------------------------------------------------------------------------------------*
'                           0                   ����I��                                                            *
'                           1                   ����I�� (��ڕs��)                                                 *
'                           -1                  �ُ�I��                                                            *
'===================================================================================================================*

Public Function funGetFixLengCut(ByVal sProccd As String, ByVal sCryNo As String, sTgetHinban As tFullHinban, _
                                 ByVal iFixCutLeng As Integer, ByVal iAllLeng As Integer, ByVal iSprFlg As Integer, _
                                 ByRef tCutBlkHinban() As typ_CutBlkHinban, ByVal iErr_Code As Integer, sErr_Msg As String) As Integer
' 1   sProccd                     String              ��      I       �H���ԍ�
' 2   sCryNo                      String              ��      I       ����w�����A���́A�����ԍ�
' 3   sTgetHinban                 String              ��      I       �_���i��
' 4   iFixCutLeng                 Integer             ��      I       ��ڕ�
' 5   iAllLeng                    Integer             ��      I       ���グ��
' 6   iSprFlg                     Integer             ��      I       �_���i�ԂŒ��(0�F�_���i�Ԓ��,1�G�z��ϐ��Œ��)
' 7   tCutBlkHinban()             typ_CutBlkHinban    ��      I/O     �ؒf��ۯ��i�ԍ\����(�z��)
' 8   iErr_Code                   Integer             ��      O       �װ����(���펞��0)
' 9   sErr_Msg                    String              ��      O       �װү���޺���

    On Error GoTo ErrorHandler
    
    '�z��J�E���g
    Dim hinban_ichi_flg             As Integer
    Dim hinban_ichi_flg1            As Integer
    
    '��ډ\��
    Dim w_FixCutLeng  As Integer
    Dim teijyaku_ok_length          As Integer              '��ڃJ�b�g�\�����ێ�
    Dim top_teijyaku, bot_teijyaku  As Integer
    
    '�g�p�l�̍ŉ��ʕ���
    Dim under_length                As String               '�ŉ������ێ��ϐ� �֐��Ăяo���p
    Dim iunder_leng                 As Integer              '���ۂ̒l
    
    '�Y��
    Dim w_i                         As Integer
    Dim w_x                         As Integer
    Dim w_y                         As Integer
    Dim indx                        As Integer
    
    '��ڔz��i�[
    Dim eCutBlkHinban()             As typ_ChangeHinb       '��ڒ�����ɂf�y�i�Ԃ�����ꍇ�ɕt������z��
    Dim wCutBlkHinban()             As typ_ChangeHinb       '�f�y�i�Ԃ��������z��
    Dim wCutBlkHinban1()            As typ_ChangeHinb       '�ŉ��ʕ��ɂ��`�F�b�N��ɒ������K�v�ȏꍇ�ɕK�v
    Dim ChangeHin()                 As typ_ChangeHinb       '��ʂ̒�ڗp�ɕ�������B�܂��ŉ��ʕ����l������������B
    Dim intHin()                    As typ_ChangeHinb       '��ʂ̒�ڂɉ������z��𓯈�i�Ԃ��ؒf�ɂďW�񂷂�B
    
    
    '�ŉ��ʕ��������ɕK�v�ϐ�
    Dim flg                         As Boolean
    Dim w_LENG1                     As Integer
    Dim w_LENG2                     As Integer
    Dim w_sa                        As Integer
    Dim w_pos1                      As Integer
    Dim w_pos2                      As Integer
    Dim cnt                         As Integer
    
    '�z��֊i�[���ɑΏ۔z��Y��
    Dim c_pos                       As Integer
    
    '�_���i�Ԃł̔z��쐬���Ɋ�z�񂪕K�v�����߂�
    Dim W_HIN1                      As Double
    
    '�f�y���f�ϐ�
    Dim w_gztop                     As Boolean
    Dim w_gztail                    As Boolean
    
    Dim lp                          As Integer
    Dim cpCutBlkHinban()            As typ_CutBlkHinban
    
    
    
    funGetFixLengCut = 0

    teijyaku_ok_length = 0
    hinban_ichi_flg = UBound(tCutBlkHinban)
    
    
    '---------------------------------------------------------------------------------------------------------------
    If iSprFlg = 1 Then
    ' --�����ǉ�-- 03/10/22  ���グ���ȏ�̒����͏����ł��Ȃ�
            ' ����������グ���ȏ�̈ʒu���Ȃ����`�F�b�N
            For lp = hinban_ichi_flg To 1 Step -1
                
                If tCutBlkHinban(lp).INPOS > iAllLeng Then
                    sErr_Msg = "TJ006"
                    iErr_Code = 1
                    funGetFixLengCut = 1
                    Exit Function
                End If
            Next
    ' --�����ǉ�-- 03/10/22
    
    ' --�����ǉ�-- 03/10/22�@�����ɕ��בւ���
            ReDim Preserve cpCutBlkHinban(hinban_ichi_flg + 1)      ' �ꎞ�R�s�[�p�̔z��
    
            For lp = 1 To hinban_ichi_flg
                Dim pos     As Integer
                Dim lp2     As Integer
                                
                pos = lp
                For lp2 = 1 To lp - 1
                    If tCutBlkHinban(lp).INPOS < cpCutBlkHinban(lp2).INPOS Then
                        pos = lp2
                        Exit For
                    End If
                Next
                
                If pos <> lp Then
                    For lp2 = pos To lp - 1
                        cpCutBlkHinban(lp2 + 1) = cpCutBlkHinban(lp2)
                    Next
                End If
                
                cpCutBlkHinban(pos) = tCutBlkHinban(lp)
            Next
    ' --�����ǉ�-- 03/10/22
    
    ' --�����ǉ�-- 03/10/17  �ŏI�ʒu���K���z��ɐݒ肳��Ă���Ƃ͌���Ȃ�
        ' �ؒf��ۯ��i�ԍ\���̂̍ŏI�f�[�^�����グ���ɓ��������H
        If cpCutBlkHinban(hinban_ichi_flg).INPOS <> iAllLeng Then
            cpCutBlkHinban(hinban_ichi_flg).LENGTH = iAllLeng       ' �������ݒ肳��Ă��Ȃ��̂Őݒ�
            hinban_ichi_flg = hinban_ichi_flg + 1                   ' �ŏI�ʒu��ǉ�
            cpCutBlkHinban(hinban_ichi_flg).INPOS = iAllLeng        '
            cpCutBlkHinban(hinban_ichi_flg).Cut = 1                 ' �J�b�g�w�肷�鎖
            cpCutBlkHinban(hinban_ichi_flg).hinban.hinban = ""      '
        End If
    
    
        Erase tCutBlkHinban                                         ' �C���[�X���Ȃ��Ɣz����Ē�`�ł��Ȃ�
        ReDim tCutBlkHinban(hinban_ichi_flg)                        ' �������z��ׂ̈��H
        
        For lp = 1 To hinban_ichi_flg                               ' �R�s�[�������߂�
            tCutBlkHinban(lp) = cpCutBlkHinban(lp)
            
            ' ���������ԂɃf�[�^���ݒ肳��Ă��Ȃ������ꍇ�ɒ��������s���ɂȂ��Ă���̂ōĐݒ�
            If lp <> hinban_ichi_flg Then
                tCutBlkHinban(lp).LENGTH = cpCutBlkHinban(lp + 1).INPOS
            End If
        Next
    End If
    ' --�����ǉ�-- 03/10/17
    '---------------------------------------------------------------------------------------------------------------
    
    
    '-------------------------------------------------------------------------------------------
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '�p�^�[���@�`�@�|�|�|�|�|���i�Ԉʒu�w��Ȃ�
    'If hinban_ichi_flg = 0 Then
    If iSprFlg = 0 Then
        '�_���i�Ԃ̎d�l�l���擾
        '---------------------------------------------------------------------------------------
        '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '�R�A�u���b�N�P�ʕۏ؃t���O���^�[�Q�b�g�i�ԂŃ`�F�b�N����
        If Check_TBCME36_DB(sTgetHinban) = False Then
            '�G���[ : �J�b�g�s�\�ł�
            sErr_Msg = "TJ001"
            iErr_Code = 1
            funGetFixLengCut = 1
            Exit Function
        End If
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        teijyaku_ok_length = iAllLeng
    Else
    '�p�^�[���@�a�@�|�|�|�|�|���i�Ԉʒu�w�肠��
     
        '----------------------------------------------------------------------------------------
        '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '�R�A�m�F�u���b�N�P�ʕۏ؃t���O
        '��ł��t���O�������Ă��Ȃ���Β�ڃJ�b�g�͏o���Ȃ�
        For indx = 1 To hinban_ichi_flg '�Ō�̂Q�O�܂ł��`�F�b�N�Ώ۔͈�
''''' --03/10/17--  �r���y�C�f�i�Ԃ��u���b�N�P�ʕۏ�t���O�`�F�b�N�ŃG���[�ɂȂ鋰�ꂪ����
'''''            '�擪�܂��͌�����y�C�f�i�Ԃ̂Ƃ��͂c�a�`�F�b�N�͍s��Ȃ�
'''''            If indx = 1 Or indx = hinban_ichi_flg Then
                If StrComp(Trim(tCutBlkHinban(indx).hinban.hinban), "Z", vbTextCompare) <> 0 And StrComp(Trim(tCutBlkHinban(indx).hinban.hinban), "G", vbTextCompare) <> 0 Then
                    '�i�Ԃ��Ȃ�������i�󔒂�������j�t���O�`�F�b�N�����Ȃ�
                    If Trim$(tCutBlkHinban(indx).hinban.hinban) <> "" Then
                        '�c�a�u���b�N�P�ʕۏ�t���O�`�F�b�N
                        If Check_TBCME36_DB(tCutBlkHinban(indx).hinban) = False Then
                            '�G���[ : �J�b�g�s�\�ł�
                            sErr_Msg = "TJ001"
                            iErr_Code = 1
                            funGetFixLengCut = 1
                            Exit Function
                        End If
                    End If
                End If
'''''            Else
'''''                '�i�Ԃ��Ȃ�������i�󔒂�������j�t���O�`�F�b�N�����Ȃ�
'''''                If Trim$(tCutBlkHinban(indx).hinban.hinban) <> "" Then
'''''                    '�c�a�u���b�N�P�ʕۏ�t���O�`�F�b�N
'''''                    If Check_TBCME36_DB(tCutBlkHinban(indx).hinban) = False Then
'''''                        '�G���[ : �J�b�g�s�\�ł�
'''''                        sErr_Msg = "TJ001"
'''''                        iErr_Code = 1
'''''                        funGetFixLengCut = 1
'''''                        Exit Function
'''''                    End If
'''''               End If
'''''            End If
''''' --03/10/17--  �r���y�C�f�i�Ԃ��u���b�N�P�ʕۏ�t���O�`�F�b�N�ŃG���[�ɂȂ鋰�ꂪ����
            '�R END ---------------------------------------------------------------------------------
            
            '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            '�S �w��i�ԃ`�F�b�N�@�����@��ڃJ�b�g�\�����Z�o����
            
            If indx > 1 And indx < hinban_ichi_flg - 1 Then
                ' �i��/�ʒu�w��@���P�̂Ƃ��̂ݏ�������
                '�@�@�y�A�f�i�Ԓ��r���݊m�F
                If StrComp(Trim(tCutBlkHinban(indx).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(indx).hinban.hinban), "G", vbTextCompare) = 0 Then
                    sErr_Msg = "TJ004"
                    iErr_Code = 1
                    funGetFixLengCut = 1
                    Exit Function
                End If
            End If
        Next indx
        
        '�S�|�A  ��ڃJ�b�g�\�Ȓ������Z�o-----------------------------------------------------
        ' �r���� ��ڃJ�b�g�\������������
        
        '�擪�̒������擾
        ' �擪���f�C�y�i�Ԃ̂Ƃ�
        If StrComp(Trim(tCutBlkHinban(1).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(1).hinban.hinban), "G", vbTextCompare) = 0 Then
          top_teijyaku = tCutBlkHinban(2).INPOS  '�Q�Ԗڂ̒������擾
        Else
          top_teijyaku = 0 '�������ɐ擪�͂O�Ƃ���
        End If
        
        '����̒������擾-----------------------------------------
        '�Ō�����@Z,�f�i�Ԃ������� ��O�i�z��ł͂Q�O�j�̒������擾
        If StrComp(Trim(tCutBlkHinban(hinban_ichi_flg - 1).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(hinban_ichi_flg - 1).hinban.hinban), "G", vbTextCompare) = 0 Then
           bot_teijyaku = tCutBlkHinban(hinban_ichi_flg - 1).INPOS  '�P�O�̒������擾 == (�����̈ʒu)
        Else
        '�y�C�f�i�Ԃł͂Ȃ�
           bot_teijyaku = tCutBlkHinban(hinban_ichi_flg).INPOS  '�Ō���̒������擾
        End If
        
        '��ډ\�����̉�
        teijyaku_ok_length = bot_teijyaku - top_teijyaku
''        teijyaku_ok_length = bot_teijyaku
        '�S�|�A END-----------------------------------------------------------------------------
        
    '�S END ---------------------------------------------------------------------------------
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    End If
    
    '----------------------------------------------------------------------------------------
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '�T�@�ŉ��ʕ��̒����̎擾
    If funCodeDBGet("SB", "TJ", "LEN", 0, " ", under_length) <> 0 Or Val(under_length) < 0 Then
        '�f�[�^�擾���G���[
        sErr_Msg = "TJ002"
        iErr_Code = 1
        funGetFixLengCut = 1
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------------------------
    '�G���[�`�F�b�N
    '��ډ\���@���@�ŉ������@�̎��G���[
    iunder_leng = Val(under_length)
     
    If teijyaku_ok_length < Val(under_length) Then
       sErr_Msg = "TJ003"
       iErr_Code = 1
       funGetFixLengCut = 1
        Exit Function
    End If
    
    '�T END ---------------------------------------------------------------------------------
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '---------------------------------------------------------------------------------------

    
    
    '�U��ڃJ�b�g�f�[�^�쐬����

    '---------------------------------------------------------------------------------------
    '�U�|�@     �ʒu�^�i�Ԏw��Ȃ��@���O�̎��̏���
    
    If hinban_ichi_flg = 0 Then
    
        Erase tCutBlkHinban
        c_pos = 1
        w_x = 0
        w_y = 0
        '��ʂ̒�ڂɉ����Ĕz����쐬����B
        w_FixCutLeng = teijyaku_ok_length
        W_HIN1 = w_FixCutLeng / iFixCutLeng
        w_i = K_fncRoundUp(W_HIN1, 0)
        For w_x = 0 To w_i
            ReDim Preserve tCutBlkHinban(c_pos)
            If w_x = w_i Then
                tCutBlkHinban(c_pos).hinban.hinban = Space(8)
                tCutBlkHinban(c_pos).INPOS = w_FixCutLeng
                tCutBlkHinban(c_pos).LENGTH = 0
            Else
                tCutBlkHinban(c_pos).hinban.hinban = sTgetHinban.hinban
                tCutBlkHinban(c_pos).INPOS = w_y
                If w_y + iFixCutLeng > w_FixCutLeng Then
                    tCutBlkHinban(c_pos).LENGTH = w_FixCutLeng
                Else
                    tCutBlkHinban(c_pos).LENGTH = w_y + iFixCutLeng
                End If
            End If
            w_y = w_y + iFixCutLeng
            c_pos = c_pos + 1
        Next
        hinban_ichi_flg = UBound(tCutBlkHinban)
    End If
    
    '�U�|�@�@�����̏I���@-----------------------------------------------------------------
    
    Erase wCutBlkHinban     '�f�y�i�Ԃ��������z��
    Erase eCutBlkHinban     '��ڒ�����ɂf�y�i�Ԃ�����ꍇ�ɕt������z��
    Erase ChangeHin         '��ʂ̒�ڗp�ɕ�������B�܂��ŉ��ʕ����l������������B
    Erase intHin
    c_pos = 1
    
    'Z,�f�i�Ԃ�ʔz��Ɋi�[���A����������ɔz��ɐݒ肷��B
    ' �擪���f�C�y�i�Ԃ̂Ƃ�
    If StrComp(Trim(tCutBlkHinban(1).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(1).hinban.hinban), "G", vbTextCompare) = 0 Then
        w_gztop = True
    Else
        w_gztop = False
    End If
    '�Ō�����@Z,�f�i�Ԃ̂Ƃ�
    If StrComp(Trim(tCutBlkHinban(hinban_ichi_flg - 1).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(hinban_ichi_flg - 1).hinban.hinban), "G", vbTextCompare) = 0 Then
        w_gztail = True
    Else
        w_gztail = False
    End If
    For w_i = 1 To hinban_ichi_flg
        'Z,�f�i�Ԃ��ȊO��ʔz��Ɋi�[���A����������ɂf�y�i�ԂƑg�ݍ��킹��
        If StrComp(Trim(tCutBlkHinban(w_i).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(w_i).hinban.hinban), "G", vbTextCompare) = 0 Then
                If w_i = hinban_ichi_flg - 1 Then
                    ReDim Preserve wCutBlkHinban(c_pos)
                    wCutBlkHinban(c_pos).hinban = tCutBlkHinban(hinban_ichi_flg).hinban
                    wCutBlkHinban(c_pos).INPOS = tCutBlkHinban(w_i).INPOS
                    wCutBlkHinban(c_pos).LENGTH = tCutBlkHinban(hinban_ichi_flg).LENGTH
                    Exit For
                End If
        Else
                ReDim Preserve wCutBlkHinban(c_pos)
                wCutBlkHinban(c_pos).hinban = tCutBlkHinban(w_i).hinban
                wCutBlkHinban(c_pos).INPOS = tCutBlkHinban(w_i).INPOS
                wCutBlkHinban(c_pos).LENGTH = tCutBlkHinban(w_i).LENGTH
                c_pos = c_pos + 1
        End If
    Next
    
    '�ŉ��ʕ��������Ɏg�p����בҔ�
    wCutBlkHinban1() = wCutBlkHinban()
    
    c_pos = 1
    w_pos1 = 0
    w_pos2 = 0
    w_sa = 0
    w_LENG1 = 0
    w_LENG2 = 0
    cnt = 0
    hinban_ichi_flg1 = UBound(wCutBlkHinban)
    '��ʂ̒�ڂɉ����Ĕz��𕪉�����B
    w_FixCutLeng = iFixCutLeng
    flg = True
    For w_i = 1 To hinban_ichi_flg1
        ReDim Preserve ChangeHin(c_pos)
        '��ʂ̒�ڂƕϐ����������Ƃ��i�J�b�g�̂͂��܂�j�𔻒f
        If w_FixCutLeng = iFixCutLeng Then
            ChangeHin(c_pos).Cut = EnumCutFlag.DoCut
        Else
            ChangeHin(c_pos).Cut = EnumCutFlag.NoCut
        End If
        '�ŉ��ʕ��������K�v�ȃJ�b�g�����邩�`�F�b�N����B����ꍇ�̓J�b�g���𒲐�
        If Not flg Then
            If w_pos1 = wCutBlkHinban(w_i).INPOS Then
                w_FixCutLeng = w_LENG1
                iFixCutLeng = w_LENG1
                cnt = cnt + 1
            Else
                If cnt = 1 Then
                    If ChangeHin(c_pos).Cut = EnumCutFlag.DoCut Then
                        w_FixCutLeng = w_LENG2
                        iFixCutLeng = w_LENG2
                        flg = True
                    End If
                End If
            End If
        End If
        '��ʂ̒�ڂɍ��킹�Ĕz��𕪉�
        If (wCutBlkHinban(w_i).LENGTH - wCutBlkHinban(w_i).INPOS) <= w_FixCutLeng Then
            ChangeHin(c_pos).hinban = wCutBlkHinban(w_i).hinban
            ChangeHin(c_pos).INPOS = wCutBlkHinban(w_i).INPOS
            ChangeHin(c_pos).LENGTH = wCutBlkHinban(w_i).LENGTH - wCutBlkHinban(w_i).INPOS
            w_FixCutLeng = w_FixCutLeng - ChangeHin(c_pos).LENGTH
            '���̃J�b�g��ؒf
            If w_FixCutLeng = 0 Then
                w_FixCutLeng = iFixCutLeng
            End If
        Else
            ChangeHin(c_pos).hinban = wCutBlkHinban(w_i).hinban
            ChangeHin(c_pos).INPOS = wCutBlkHinban(w_i).INPOS
            ChangeHin(c_pos).LENGTH = w_FixCutLeng
            wCutBlkHinban(w_i).INPOS = wCutBlkHinban(w_i).INPOS + w_FixCutLeng
            w_FixCutLeng = iFixCutLeng
            w_i = w_i - 1
        End If
        '�ŏI�J�b�g�͋����I�ɐؒf
        If w_i = hinban_ichi_flg1 Then
            ChangeHin(c_pos).Cut = EnumCutFlag.DoCut
           '�ŏI�J�b�g���ŉ��ʕ���菬�����ꍇ�Ĕz�񒲐����K�v���`�F�b�N
            flg = True
            For w_x = c_pos - 1 To 1 Step -1
                If flg Then
                    If ChangeHin(w_x).Cut = EnumCutFlag.DoCut Then
                        If ChangeHin(c_pos).INPOS - ChangeHin(w_x).INPOS < Val(under_length) Then
                            w_sa = Val(under_length) - (ChangeHin(c_pos).INPOS - ChangeHin(w_x).INPOS)
                            w_LENG2 = Val(under_length)
                            w_pos2 = ChangeHin(w_x).INPOS
                            flg = False
                        Else
                            Exit For
                        End If
                    End If
                Else
                    If ChangeHin(w_x).Cut = EnumCutFlag.DoCut Then
                            w_pos1 = ChangeHin(w_x).INPOS
                            w_LENG1 = (w_pos2 - ChangeHin(w_x).INPOS) - w_sa
                            Exit For
                    End If
                End If
            Next
            '�ŉ��ʕ��������K�v�ȈׁA�ēx�z��쐬�i�e�ϐ������ݒ�ɖ߂��j
            If Not flg Then
                Erase ChangeHin
                w_FixCutLeng = iFixCutLeng
                wCutBlkHinban() = wCutBlkHinban1()
                w_i = 0
                c_pos = 0
            End If
        End If
        c_pos = c_pos + 1
    Next
    
    c_pos = 0
    '��L�̉�ʂ̒�ڂɉ������z��𓯈�i�Ԃ��ؒf�ɂďW�񂷂�B
    For w_i = 1 To UBound(ChangeHin)
        If ChangeHin(w_i).Cut = EnumCutFlag.DoCut Then
            c_pos = c_pos + 1
             ReDim Preserve intHin(c_pos)
             intHin(c_pos).hinban = ChangeHin(w_i).hinban
            intHin(c_pos).INPOS = ChangeHin(w_i).INPOS
             intHin(c_pos).Cut = ChangeHin(w_i).Cut
        Else
            If Trim$(ChangeHin(w_i - 1).hinban.hinban) <> Trim$(ChangeHin(w_i).hinban.hinban) Then
                 c_pos = c_pos + 1
                 ReDim Preserve intHin(c_pos)
                 intHin(c_pos).hinban = ChangeHin(w_i).hinban
                 intHin(c_pos).INPOS = ChangeHin(w_i).INPOS
                 intHin(c_pos).Cut = ChangeHin(w_i).Cut
            End If
        End If
    Next
    
    '�ŉ��ʕ��`�F�b�N���s���B�ŉ����ȉ��̃J�b�g��������ꍇ�̓G���[
    c_pos = 0
    For w_i = 1 To UBound(intHin)
        '�ؒf�J�b�g��
        If intHin(w_i).Cut = EnumCutFlag.DoCut Then
            For w_x = w_i + 1 To UBound(intHin)
                '�ؒf�J�b�g��
                If intHin(w_x).Cut = EnumCutFlag.DoCut Then
                    If intHin(w_x).INPOS - intHin(w_i).INPOS < Val(under_length) Then
                        sErr_Msg = "TJ005"
                        iErr_Code = 1
                        funGetFixLengCut = 1
                        Exit Function
                    Else
                        Exit For
                    End If
                End If
            Next
        End If
    Next
    
    'Z,G�i�Ԃ�����
    c_pos = 0
    If w_gztop = True Then
        c_pos = c_pos + 1
        ReDim Preserve eCutBlkHinban(c_pos)
        eCutBlkHinban(c_pos).hinban = tCutBlkHinban(1).hinban
        eCutBlkHinban(c_pos).INPOS = tCutBlkHinban(1).INPOS
        eCutBlkHinban(c_pos).Cut = tCutBlkHinban(1).Cut
    End If
    For w_i = 1 To UBound(intHin)
        c_pos = c_pos + 1
        ReDim Preserve eCutBlkHinban(c_pos)
        eCutBlkHinban(c_pos).hinban = intHin(w_i).hinban
        eCutBlkHinban(c_pos).INPOS = intHin(w_i).INPOS
        eCutBlkHinban(c_pos).Cut = intHin(w_i).Cut
    Next
    If w_gztail = True Then
        ReDim Preserve eCutBlkHinban(c_pos)
        eCutBlkHinban(c_pos).hinban = tCutBlkHinban(hinban_ichi_flg - 1).hinban
        eCutBlkHinban(c_pos).INPOS = tCutBlkHinban(hinban_ichi_flg - 1).INPOS
        eCutBlkHinban(c_pos).Cut = tCutBlkHinban(hinban_ichi_flg - 1).Cut
        c_pos = c_pos + 1
        ReDim Preserve eCutBlkHinban(c_pos)
        eCutBlkHinban(c_pos).hinban = tCutBlkHinban(hinban_ichi_flg).hinban
        eCutBlkHinban(c_pos).INPOS = tCutBlkHinban(hinban_ichi_flg).INPOS
        eCutBlkHinban(c_pos).Cut = tCutBlkHinban(hinban_ichi_flg).Cut
    End If
    
    '�k�d�m�f�s�g�𒲐�
    c_pos = 0
    For w_i = 1 To UBound(eCutBlkHinban) - 1
        eCutBlkHinban(w_i).LENGTH = eCutBlkHinban(w_i + 1).INPOS - eCutBlkHinban(w_i).INPOS
    Next
    
    '�V�@���ʃZ�b�g���ďo���֕Ԃ��B
    tCutBlkHinban() = eCutBlkHinban()

    
    
Error:
    Exit Function
    
ErrorHandler:
    funGetFixLengCut = -1
    GoTo Error

End Function
' funGetFixLengCut  END OF FUNCTION          ---------------------------------------------------------------------------------------------------

'====================================================================================================================
Rem                                                                                                                 *
Rem �E�u���b�N�P�ʕۏ� �e�[�u�����Q�Ƃ��u���b�N�P�ʕۏ؃t���O���`�F�b�N����t�@���N�V�����@                         *
'====================================================================================================================
'  �Q�ƃe�[�u��                         TBCME036                                                                    *
'  ���ږ��@�@�@��ۯ��P�ʕۏ��׸�        BLOCKHFLAG  0: �w�肳�ꂽ�A�_���i�Ԃ̎d�l�l���擾����B                     *
'                                                   1: �w�肳��Ă���S�i�Ԃ̎d�l�l���擾����B�i�z��쐬�j         *
'-------------------------------------------------------------------------------------------------------------------*
'                                                                                                                   *
' �����Q                                                                                                            *
'�@��P�����@�@�i��   :�t���O���Q�Ƃ���i��                                                                         *
'                                                                                                                   *
'-------------------------------------------------------------------------------------------------------------------*
'   �߂�l                                                                                                          *
'   ����I��               :TRUE       �u���b�N�t���O��0�i��ډ\�j                                              *
'                           :FALSE      �u���b�N�t���O��1 �i��ڕs�\�j                                           *
'-------------------------------------------------------------------------------------------------------------------*
'                           0                   ����I��                                                            *
'                           1                   ����I�� (��ڕs��)                                                 *
'                           -1                  �ُ�I��                                                            *
'===================================================================================================================*

Function Check_TBCME36_DB(Fhinban As tFullHinban) As Boolean
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim recCnt      As Long         '���R�[�h��
    Dim i           As Long
    Dim dbIsMine    As Boolean

    On Error GoTo ErrorHandler

    ''SQL��g�ݗ��Ă�
    sql = "Select BLOCKHFLAG From TBCME036"
    sql = sql & " where HINBAN   = '" & Fhinban.hinban & "' "
    sql = sql & "   and MNOREVNO =  " & Fhinban.mnorevno
    sql = sql & "   and FACTORY  = '" & Fhinban.factory & "' "
    sql = sql & "   and OPECOND  = '" & Fhinban.opecond & "' "

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        Check_TBCME36_DB = False
        Exit Function
    End If

    If rs.RecordCount <> 0 Then
        If rs.Fields(0) = "0" Then
                Check_TBCME36_DB = True
         Else
                Check_TBCME36_DB = False
        End If
    Else
         Check_TBCME36_DB = False
    End If
        
Error:
    Exit Function
    
ErrorHandler:
    Check_TBCME36_DB = False
    GoTo Error

End Function
'============================================================
'
'�@�\            :�؂�グ
'
'������          :�@�؂�グ���鐔�l
'�@�@�@           �A������H�ʂ́H��
'
'�߂�l          :�@�؂�グ��̐��l
'
'�@�\����        :���l�̐؂�グ���Ɏg�p����B
'
'���l            :
'
'============================================================
Public Function K_fncRoundUp(Mdbl_Su As Double, Mint_keta As Integer) As Currency
    
    Dim Mcur_work   As Currency
    Dim Mdbl_ret    As Double
    
    On Error GoTo ErrorHandler
    
    K_fncRoundUp = True
    
    Mcur_work = Mdbl_Su
    
    Mdbl_ret = 10 ^ Abs(Mint_keta)
    
    If Mint_keta > 0 Then
        K_fncRoundUp = Int(Mcur_work * Mdbl_ret + 0.9999) / Mdbl_ret
    Else
        K_fncRoundUp = Int(Mcur_work / Mdbl_ret + 0.9999) * Mdbl_ret
    End If
    
Error:
    Exit Function
    
ErrorHandler:
    K_fncRoundUp = False
    GoTo Error

End Function

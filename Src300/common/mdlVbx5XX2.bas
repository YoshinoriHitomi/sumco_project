Attribute VB_Name = "mdlVbx5xx2"
' @(h) mdlVbx5XX2.BAS              ver 1.00 ( '00.01.06 �a���)
' @(s)
Option Explicit
''���o�����ݒ�E���ݒ�t���O
''�ݒ�ς�=True ���ݒ�=False
Public gbFlgVbx5xx2 As Boolean
''���C��SQL�ƌ������鍀�ڂ��w�肷��B
''���������ԍ��Fxtalc2,3�� ���㌋���ԍ��Fowxtalc2 �i�ԁFhinbc2
Public gsKeyVbx5xx2 As String
'' ����WHERE�������O���[�o���ϐ�
'' �e�[�u�����FJKB=�������� JKC=���㌋�� JKD=�����w�� JKE=�d�l
Public gsSqlWhereVbx5xx2 As String

'�e�[�u���A�N�Z�X�t���O
'���ڂɂ�茋������e�[�u����I������B
'�g�p=True,���g�p=False �iCheckF4Vbx5XX2�Őݒ�j
Private bXODC1 As Boolean   '���㌋��
Private bXODC2 As Boolean   '9/13 Yam
Private bXODE2 As Boolean   '�����w��
Private bSIYO1 As Boolean   '�d�l�P
Private bSIYO2 As Boolean   '�d�l�Q
Private bSIYO3 As Boolean   '�d�l�R

''VBX5041�[���ݒ�E���ݒ�t���O
''�ݒ�ς�=True ���ݒ�=False
Public gbFlgVbx5040Nouki As Boolean
'�ϊ��O�̋@��R�[�h�̕ۑ�(11/17 Yam�ǉ��j
Public kisyuNm As String

Type CDNAMEDAT  ''�R�[�h�E���O
    Cd As String
    Nm As String
End Type

' @(f)
' �@�\      : �L�[���䏈��
' �Ԃ�l    : �Ȃ�
' ����      : KeyCode   -   �L�[�R�[�h
' �@�\����  : �L�[�R�[�h�ɂ���ď�����U�蕪���鏈�����s��
' ���l      : ��ʏ�ԃt���O     0:�������
'                               1:�m�F���s
'                               2:�o�^���s
'
Public Sub KeyActionVbx5XX2(KeyCode As Integer)

gbFlgVbx5040Nouki = False   ''2000/06/07�C��

    With frmSub
        ''�R�}���h�{�^���@�\�U����
        Select Case KeyCode
        Case vbKeyF3
            '''�L�����Z��
            If .cmdF(3).Enabled = False Then Exit Sub
            ''��ʏ�����
            Call InitVbx5XX2(True)
            ''���o�����ݒ�OFF
            gbFlgVbx5xx2 = False
            ''���C����ʕ��A
            frmMain.Show
            frmSub.Hide
        Case vbKeyF4
            '''�C��
            If .cmdF(4).Enabled = False Then Exit Sub
            ''���o��ʐ�p�⍇�����쐬
            If vbKeyActionF4Vbx5XX2() Then
                ''��ʏ�����
                Call InitVbx5XX2(False)
                ''���o�����ݒ�ON
                gbFlgVbx5xx2 = True
                ''���C����ʕ��A
                frmMain.Show
                frmSub.Hide
            End If
        End Select
    End With
End Sub

' @(f)
' �@�\      : ���o����WHERE���쐬�iMAIN�j
' �Ԃ�l    : �Ȃ�
' ����      : �Ȃ�
' �@�\����  : ���o������ʂŐݒ肵�����ڂ̏��������쐬����B
' ���l      :

Private Function vbKeyActionF4Vbx5XX2()
    Dim i           As Integer      ''���[�v�J�E���^
    Dim sWk         As String       ''��Ɨ̈�
    Dim iWild       As Integer      ''���C���h�J�[�h�g�p�t���O
    vbKeyActionF4Vbx5XX2 = False
    gbFlgVbx5040Nouki = False
    ''���͍��ڂ̃`�F�b�N
    If CheckF4Vbx5XX2() = False Then
        Exit Function
    End If
    
    ''���o������� �⍇�����̍쐬
    ''ex)�쐬�C���[�W
    ''      ,�������� JKB, ���㌋�� JKC, �����w�� JKD, �d�l JKE
    ''      WHERE �i���C�����SQL�jJKA.�����L�[ = JKB.�Ή�����L�[
    ''      AND JKB.���㌋�� = JKC.���㌋��
    ''      AND JKB.�i�� = JKD.�i��
    ''      AND JKB.�i�� = JKE.�i��
    ''      AND �ȉ��I�����ڂɂ���
    ''���g�p����e�[�u���̋L�q
    ''  ���������e�[�u���͕K����������B
    ''  ���������e�[�u��������㌋���E�����w���A�d�l�e�[�u������������B
    If bXODC2 Then   '9/13 Yam
        gsSqlWhereVbx5xx2 = ", xodc2 JKB "          ''������������(�ʖ��FJKB)
    Else
        gsSqlWhereVbx5xx2 = " "                     ''������������(�ʖ��FJKB)
    End If
    If bXODC1 Then
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & ", xodc1 JKC"     ''���㌋��(�ʖ��FJKC)
    End If
    'If bXODE2 Then
    '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & ", xode2 JKD"     ''�����w������(�ʖ��FJKD)
    'End If
    If bSIYO1 Then
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & ", sods1 JKE"     ''�d�l�P����(�ʖ��FJKE)
    End If
    If bSIYO2 Then
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & ", sods2_es JKF"  ''�d�l�P����(�ʖ��FJKE)
    End If
    If bSIYO3 Then
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & ", sods2_pr JKG"  ''�d�l�P����(�ʖ��FJKE)
    End If
        
    ''���L�[�������ɋL�q
    ''  ���C�����SQL�Ǝg�p�e�[�u���̌����������쐬����B
    ''  ��������(JKB)�Ƃ͕K����������B
    ''���C��SQL����������(JKB).�i���C��SQL���ێ�����L�[�Ɉˑ��j
    ''���������Ō�������ꍇ
    If bXODC2 Then  '9/13 Yam
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " where JKA.ck1 = JKB." & gsKeyVbx5xx2
    Else
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " where JKA.ck1 = JKA.ck1 "
    End If
    If bXODC1 Then
        ''��������(JKB).���㌋���ԍ�=���㌋��.���㌋���ԍ�
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.owxtalc2 = JKC.xtalc1"
    End If
    'If bXODE2 Then
        ''��������(JKB).���㌋���ԍ�=�����w��.�i��
        ''  2000/06/19  �C���J�n    ���������e�[�u���̕i�Ԃƌ������Ȃ��ōH�����т�
        ''  �i�Ԃƌ������邽�ߏC���i���r�W���������������) ����  ���c
'        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.hinbc2 || JKB.hinbrc2 = JKD.hinbe2 || JKD.hinbre2"
        ''02/14/2000 �����w��(XODE2)�̈�ʉ��̂��߂ɒǉ�
    '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK2   =   JKD.hinbe2"
    '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK3   =   JKD.hinbre2"
    '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.knnoc2 = JKD.knnoe2"
        ''  2000/06/19  �C�������܂�    ���������e�[�u���̕i�Ԃƌ������Ȃ��ōH�����т�
        ''  �i�Ԃƌ������邽�ߏC���i���r�W�����A���Ԃ����������) ����  ���c
    'End If
    If bSIYO1 Then
        ''��������(JKB).���㌋���ԍ�=�d�l.�i��
'        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.hinbc2 || JKB.hinbrc2 = JKE.hinbc3 || JKE.hinbrc3"
        ''  2000/06/19  �C���J�n    ���������e�[�u���̕i�Ԃƌ������Ȃ��ōH�����т�
        ''  �i�Ԃƌ������邽�ߏC���i���r�W���������������) ����  ���c
'        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.hinbc2 = JKE.hinbc3 and JKB.hinbrc2 = JKE.hinbrc3"
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK2   = JKE.specnos1 "
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK3   = JKE.specnors1"
        ''  2000/06/19  �C�������܂�    ���������e�[�u���̕i�Ԃƌ������Ȃ��ōH�����т�
        ''  �i�Ԃƌ������邽�ߏC���i���r�W���������������) ����  ���c
    End If
    If bSIYO2 Then
        ''��������(JKB).���㌋���ԍ�=�d�l.�i��
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK2   = JKF.specnos2 "
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK3   = JKF.specnors2"
    End If
    If bSIYO3 Then
        ''��������(JKB).���㌋���ԍ�=�d�l.�i��
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK2   = JKG.specnos2 "
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK3   = JKG.specnors2"
    End If
    
    ''�����o��ʍ��ڂ̏��������L�q
    With frmVBX5XX2
        ''�i��(��������)
        If .optHinban(0).Value Then
            ''��v
            If Trim(.txtHinban(0).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.ck2 || JKA.ck3 >= '" & Trim(UCase(.txtHinban(0).Text)) & "'"
            End If
            If Trim(.txtHinban(1).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.ck2 || JKA.ck3 <= '" & Trim(UCase(.txtHinban(1).Text)) & "'"
            End If
        Else
            ''�s��v
            If Trim(.txtHinban(0).Text) <> "" And Trim(.txtHinban(1).Text) <> "" Then
                    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKA.ck2 || JKA.ck3 < '" & Trim(UCase(.txtHinban(0).Text)) & "'"
                    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " or JKA.ck2 || JKA.ck3 > '" & Trim(UCase(.txtHinban(1).Text)) & "')"
            Else
                If Trim(.txtHinban(0).Text) <> "" Then
                    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.ck2 || JKA.ck3 <= '" & Trim(UCase(.txtHinban(0).Text)) & "'"
                End If
                If Trim(.txtHinban(1).Text) <> "" Then
                    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.ck2 || JKA.ck3 >= '" & Trim(UCase(.txtHinban(1).Text)) & "'"
                End If
            End If
        End If
        ''�@��(��������)
        If Trim(.txtKisy.Text) <> "" Then
            If .optKisy(0).Value Then
                ''��v
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and SUBSTR(JKB.kisyuc2,1,2) = '" & Trim(.txtKisy.Text) & "'"
            Else
                ''�s��v
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and SUBSTR(JKB.kisyuc2,1,2) != '" & Trim(.txtKisy.Text) & "'"
            End If
        End If
        
        .txtKisy.Text = kisyuNm   '11/17 �ǉ�(Yam)
        
        ''������@(��������)
        sWk = ""
        For i = 0 To 2
            If Trim(.txtHikiageX(i).Text) <> "" Then
                sWk = sWk & "'" & Trim(.txtHikiageX(i).Text) & "',"
            End If
        Next
        If sWk <> "" Then
            sWk = Mid(sWk, 1, Len(sWk) - 1) '�Ō�̃J���}���Ƃ�B
            If .optHikiageX(0).Value Then
                ''��v
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.pumethc2 in(" & sWk & ")"
            Else
                ''�s��v
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.pumethc2 not in(" & sWk & ")"
            End If
        End If
        ''PG-ID(���㌋��)
        .txtPgid.Text = UCase(.txtPgid.Text)   'Yam�ǉ�
        If Trim(.txtPgid.Text) <> "" Then
            sWk = Trim(.txtPgid.Text)
            ''���C���h�J�[�h����[?]��[_]�ɕϊ�����(1�������C���h�J�[�h)
            Do
                iWild = InStr(sWk, "?")
                If iWild = 0 Then
                    Exit Do
                Else
                    sWk = Mid(sWk, 1, iWild - 1) & "_" & Mid(sWk, iWild + 1)
                End If
            Loop
            ''8���ɖ����Ȃ��ꍇ�͍Ō��[%]��t����
            If Len(sWk) < 8 Then
                If Right(sWk, 1) <> "%" Then
                    sWk = sWk & "%"
                End If
            End If
            If .optPgid(0).Value Then
                ''��v
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKC.pgidc1 like '" & sWk & "'"
            Else
                ''�s��v
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKC.pgidc1 not like '" & sWk & "'"
            End If
        End If
        ''���i�敪(�d�l)  1/22 Yam �C��
        If Trim(.txtSeizoKbn.Text) <> "" Then
                '4:���̑��̏ꍇ
            If Trim(.txtSeizoKbn.Text) = "9" Then
                If .optSeizoKbn(0).Value Then
                    ''��v
                    For i = 1 To 3
                        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.comprdgrs1 != '" & Right("00" & Trim(i), 2) & "'"
                    Next
                Else
                    ''�s��v
                        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and ((JKE.comprdgrs1 = '01') or (JKE.comprdgrs1 = '02') or (JKE.comprdgrs1 = '03'))"
                End If
                '1,2,3:���̑��ȊO�̏ꍇ
            Else
                If .optSeizoKbn(0).Value Then
                    ''��v
                        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.comprdgrs1 = '" & Right("00" & Trim(.txtSeizoKbn.Text), 2) & "'"
                Else
                    ''�s��v
                    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.comprdgrs1 != '" & Right("00" & Trim(.txtSeizoKbn.Text), 2) & "'"
                End If
            End If
        End If
        ''�g�p�ړI(�d�l)
        sWk = ""
        For i = 0 To 5
            If Trim(.txtMokuteki(i).Text) <> "" Then
                sWk = sWk & "'" & Trim(.txtMokuteki(i).Text) & "'" & ","
            End If
        Next
        If sWk <> "" Then
            sWk = Mid(sWk, 1, Len(sWk) - 1) '�Ō�̃J���}���Ƃ�B
            If .optMokuteki(0).Value Then
                ''��v
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.comcususes1 in(" & sWk & ")"
            Else
                ''�s��v
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.comcususes1 not in(" & sWk & ")"
            End If
        End If
        ''����(�����w��) ��kuro��������DB��萻���w��DB�ɕύX��
        'sWk = ""
        'For i = 0 To 1
        '    If Trim(.txtMukaisaki(i).Text) <> "" Then
        '        sWk = sWk & "'" & Trim(.txtMukaisaki(i).Text) & "',"
        '    End If
        'Next
        'If sWk <> "" Then
        '    sWk = Mid(sWk, 1, Len(sWk) - 1) '�Ō�̃J���}���Ƃ�B
        '    If .optMukaisaki(0).Value Then
        '        ''��v
        '        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKD.swplace2 in(" & sWk & ")"
        '    Else
        '        ''�s��v
        '        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKD.swplace2 not in(" & sWk & ")"
        '    End If
        'End If
        ''�[��(�����w��)
        'If Trim(.txtNoki(0).Text) <> "" Then
        '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and TO_CHAR(JKD.snyye2,'FM0000') || TO_CHAR(JKD.snmme2,'FM00') || TO_CHAR(JKD.sndde2,'FM00') >= '"
        '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & DateChange(Trim(.txtNoki(0).Text)) & "'"
        'End If
        'If Trim(.txtNoki(1).Text) <> "" Then
        '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and TO_CHAR(JKD.snyye2,'FM0000') || TO_CHAR(JKD.snmme2,'FM00') || TO_CHAR(JKD.sndde2,'FM00') <= '"
        '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & DateChange(Trim(.txtNoki(1).Text)) & "'"
        'End If
        ''���@(��������)
        If Trim(.txtGoki(0).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.owxtalc2 >= '" & Trim(.txtGoki(0).Text) & "'"
        End If
        If Trim(.txtGoki(1).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.owxtalc2 <= '" & Trim(.txtGoki(1).Text) & "999999999'"
        End If
        ''�i��敪(��������)
        If Trim(.txtKakuage.Text) <> "" Then
        ' 1/23 Yam�C��   gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.laupc2 != '" & Trim(.txtKakuage.Text) & "ZZZZZZZZZ'"
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.laupc2 != '" & Trim(.txtKakuage.Text) & "'"
        End If
        ''���a�敪(�d�l)
        If Trim(.txtChokkei(0).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.commxdiadvs1 >= '" & Trim(.txtChokkei(0).Text) & "'"
        End If
        If Trim(.txtChokkei(1).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.commxdiadvs1 <= '" & Trim(.txtChokkei(1).Text) & "'"
        End If
        ''�`���^(�d�l)
        '////Chihi 11/14 �ǉ�///
        .txtDendo.Text = UCase(.txtDendo.Text)
        If Trim(.txtDendo.Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKG.mxtyps2 = '" & Trim(.txtDendo.Text) & "'"
        End If
        ''�h�[�p���g(�d�l)
        '/// Chihi 11/13�ǉ�
        .txtDoba.Text = UCase(.txtDoba.Text)
        If Trim(.txtDoba.Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKG.mxdops2 = '" & Trim(.txtDoba.Text) & "'"
        End If
        ''����(�d�l)
        If Trim(.txtHoui.Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKG.axaxiss2 = '" & Trim(.txtHoui.Text) & "'"
        End If
        ''��R��(�d�l�j
        If Trim(.txtTeikoKbn) <> "" Then
            If Trim(.txtTeikouritsu(0).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.rsxtalls2 >= " & Trim(.txtTeikouritsu(0).Text)
            End If
            If Trim(.txtTeikouritsu(1).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.rsxtalls2 <= " & Trim(.txtTeikouritsu(1).Text)
            End If
        Else
            If Trim(.txtTeikouritsu(0).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.rsxtalus2 >= " & Trim(.txtTeikouritsu(0).Text)
            End If
            If Trim(.txtTeikouritsu(1).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.rsxtalus2 <= " & Trim(.txtTeikouritsu(1).Text)
            End If
        End If
        ''��R(�����W)(�d�l)
        If Trim(.txtTeikou(0).Text) <> "" Then
            ''���X���O��Rmin(JKE.teikou_s)��null�̎��̃��R�[�h�͒��o�ł��܂���B
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.teikou_u / JKF.teikou_s) >= " & Trim(.txtTeikou(0).Text)
            ''���X���O��Rmin(JKE.teikou_s)��0�̎��ɂ͏�����0�ɂȂ�A�G���[���Ԃ��Ă��܂��B(����0��1�ɒu��������ꍇ��)
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.teikou_u / DECODE(JKF.teikou_s,0,1,JKF.teikou_s)) >= " & Trim(.txtTeikou(0).Text)
            ''���X���O��Rmin(JKE.teikou_s)��0�̎��ɂ͏�����0�ɂȂ�A�G���[���Ԃ��Ă��܂��B(����0��null�ɒu��������ꍇ��)
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.rsxtalus2 / DECODE(JKF.rsxtalls2,0,null,JKF.rsxtalls2)) >= " & Trim(.txtTeikou(0).Text)

        End If
        If Trim(.txtTeikou(1).Text) <> "" Then
            ''���X���O��Rmin(JKE.teikou_s)��null�̎��̃��R�[�h�͒��o�ł��܂���B
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.teikou_u / JKF.teikou_s) <= " & Trim(.txtTeikou(1).Text)
            ''���X���O��Rmin(JKE.teikou_s)��0�̎��ɂ͏�����0�ɂȂ�A�G���[���Ԃ��Ă��܂��B(����0��1�ɒu��������ꍇ��)
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.teikou_u / DECODE(JKF.teikou_s,0,1,JKF.teikou_s)) <= " & Trim(.txtTeikou(1).Text)
            ''���X���O��Rmin(JKE.teikou_s)��0�̎��ɂ͏�����0�ɂȂ�A�G���[���Ԃ��Ă��܂��B(����0��null�ɒu��������ꍇ��)
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.rsxtalus2 / DECODE(JKF.rsxtalls2,0,null,JKF.rsxtalls2)) <= " & Trim(.txtTeikou(1).Text)
        End If
        ''�_�f�Z�x(�d�l�j
        If Trim(.txtSansoKbn) <> "" Then
            If Trim(.txtSanso(0).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oislgls2 >= " & Trim(.txtSanso(0).Text)
            End If
            If Trim(.txtSanso(1).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oislgls2 <= " & Trim(.txtSanso(1).Text)
            End If
        Else
            If Trim(.txtSanso(0).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oislgus2 >= " & Trim(.txtSanso(0).Text)
            End If
            If Trim(.txtSanso(1).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oislgus2 <= " & Trim(.txtSanso(1).Text)
            End If
        End If
        ''Oi(�����W)(�d�l)
        If Trim(.txtOi(0).Text) <> "" Then
            ''���X���O[Oi]min(JKF.teikou_s)��null�̎��̃��R�[�h�͒��o�ł��܂���B
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.sanso_u / JKF.sanso_s) >= " & Trim(.txtOi(0).Text)
            ''���X���O[Oi]min(JKF.teikou_s)��0�̎��ɂ͏�����0�ɂȂ�A�G���[���Ԃ��Ă��܂��B(����0��1�ɒu��������ꍇ��)
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.sanso_u / DECODE(JKF.sanso_s,0,1,JKF.sanso_s)) >= " & Trim(.txtOi(0).Text)
            ''���X���O[Oi]min(JKF.teikou_s)��0�̎��ɂ͏�����0�ɂȂ�A�G���[���Ԃ��Ă��܂��B(����0��null�ɒu��������ꍇ��)
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.oislgus2 / DECODE(JKF.oislgls2,0,null,JKF.oislgls2)) >= " & Trim(.txtOi(0).Text)
        End If
        If Trim(.txtOi(1).Text) <> "" Then
            ''���X���O[Oi]min(JKF.teikou_s)��null�̎��̃��R�[�h�͒��o�ł��܂���B
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.sanso_u / JKF.sanso_s) <= " & Trim(.txtOi(1).Text)
            ''���X���O[Oi]min(JKF.teikou_s)��0�̎��ɂ͏�����0�ɂȂ�A�G���[���Ԃ��Ă��܂��B(����0��1�ɒu��������ꍇ��)
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.sanso_u / DECODE(JKF.sanso_s,0,1,JKF.sanso_s)) <= " & Trim(.txtOi(1).Text)
            ''���X���O[Oi]min(JKF.teikou_s)��0�̎��ɂ͏�����0�ɂȂ�A�G���[���Ԃ��Ă��܂��B(����0��null�ɒu��������ꍇ��)
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.oislgus2 / DECODE(JKF.oislgls2,0,null,JKF.oislgls2)) <= " & Trim(.txtOi(1).Text)
        End If
        ''ORG(�d�l)
        If Trim(.txtOrg(0).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oiorgs2 >= " & Trim(.txtOrg(0).Text)
        End If
        If Trim(.txtOrg(1).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oiorgs2 <= " & Trim(.txtOrg(1).Text)
        End If
    End With
    
    vbKeyActionF4Vbx5XX2 = True
End Function

' @(f)
' �@�\      : �L�[�`�F�b�N����
' �Ԃ�l    : ����=True,�ُ�=False
' ����      : �Ȃ�
' �@�\����  : ���͍��ڃ`�F�b�N
' ���l      :
'
Private Function CheckF4Vbx5XX2() As Boolean
    Dim i           As Integer      ''���[�v�J�E���^
    Dim akisy As String

    CheckF4Vbx5XX2 = False
    bXODC1 = False
    bXODC2 = False '9/13 Yam
    bXODE2 = False
    bSIYO1 = False
    bSIYO2 = False
    bSIYO3 = False
    With frmSub
        ''�i�ԃ`�F�b�N
        If Trim(.txtHinban(0).Text) <> "" And Trim(.txtHinban(1).Text) <> "" Then
            Call FillUpString(.txtHinban(0), "0")
            Call FillUpString(.txtHinban(1), "9")
            If Val(.txtHinban(0).Text) > Val(.txtHinban(1).Text) Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtHinban(0), RED_CTL)
                Call CtrlEnabled(.txtHinban(1), RED_CTL)
                Exit Function
            End If
            '���������e�[�u�����ڂ̂��߃e�[�u���t���O�͗��ĂȂ��B
        End If
        
        ''�@��`�F�b�N            '11/17�@Yam�ǉ�
        .txtKisy.Text = UCase(Trim(.txtKisy.Text))
        kisyuNm = .txtKisy.Text
        If Trim(.txtKisy.Text) <> "" Then
            If GetkisyNo(Trim(.txtKisy.Text), akisy) = False Then
                Exit Function
            End If
            .txtKisy.Text = akisy
            bXODC2 = True           '9/13 Yam
        End If
        ''������@�`�F�b�N
        For i = 0 To 2
            If Trim(.txtHikiageX(i).Text) <> "" Then
            bXODC2 = True           '9/13 Yam
            End If
        Next
        ''PG-ID�`�F�b�N
        If Trim(.txtPgid.Text) <> "" Then
            bXODC1 = True           '���㌋��ON
            bXODC2 = True           '9/13 Yam
        End If
        ''���i�敪�`�F�b�N
        If Trim(.txtSeizoKbn.Text) <> "" Then
            If (Val(.txtSeizoKbn.Text) < 1) Or (Val(.txtSeizoKbn.Text) > 4) _
            And (Val(.txtSeizoKbn.Text) <> 9) Or (IsNumeric(.txtSeizoKbn.Text) = False) Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtSeizoKbn, RED_CTL)
                Exit Function
            End If
            bSIYO1 = True            '�d�lON
        End If
        ''�g�p�ړI�`�F�b�N
        For i = 0 To 5
            If Trim(.txtMokuteki(i).Text) <> "" Then
                '���Ԃɓ��͂���Ă��邱��
                If i > 0 Then
                    If Trim(.txtMokuteki(i - 1).Text) = "" Then
                        Call MsgOut(50, "", ERR_DISP)
                        Call CtrlEnabled(.txtMokuteki(i - 1), RED_CTL)
                        Exit Function
                    End If
                End If
                bSIYO1 = True        '�d�lON
            End If
        Next
        ''����`�F�b�N
        'For i = 0 To 1
        '    If Trim(.txtMukaisaki(i).Text) <> "" Then
        '        '���Ԃɓ��͂���Ă��邱��
        '        If i > 0 Then
        '            If Trim(.txtMukaisaki(i - 1).Text) = "" Then
        '                Call MsgOut(50, "", ERR_DISP)
        '                Call CtrlEnabled(.txtMukaisaki(i - 1), RED_CTL)
        '                Exit Function
        '            End If
        '        End If
        '        '�����w��ON kuro�ǉ�
        '        If Len(Trim(.txtMukaisaki(i).Text)) <> 0 Then
        '            bXODE2 = True
        '        End If
        '    End If
        'Next
        ''�[���`�F�b�N
        'If Trim(.txtNoki(0).Text) <> "" Or Trim(.txtNoki(1).Text) <> "" Then
        '    '�������͂���Ă����ꍇ���𖄂߂�
        '    If Len(Trim(.txtNoki(0).Text)) <> 0 Then
        '        FillUpString .txtNoki(0), "0"
        '    ElseIf Len(Trim(.txtNoki(1).Text)) <> 0 Then
        '        FillUpString .txtNoki(1), "9"
        '    End If
            '���ԃ`�F�b�N���o�����p
        '    If KikanCheckVbx5XX2(.txtNoki(0), .txtNoki(1)) = False Then
        '        Exit Function
        '    End If
        '    bXODE2 = True            '�����w��ON
        '    gbFlgVbx5040Nouki = True '�[���ݒ�ς�
        'End If
        ''���@�`�F�b�N
        If Trim(.txtGoki(0).Text) <> "" Or Trim(.txtGoki(1).Text) <> "" Then
            If Trim(.txtGoki(0).Text) > Trim(.txtGoki(1).Text) Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtGoki(0), RED_CTL)
                Call CtrlEnabled(.txtGoki(1), RED_CTL)
                Exit Function
            End If
            bXODC2 = True           '9/13 Yam
        End If
        ''�i��敪�`�F�b�N
        If Trim(.txtKakuage.Text) <> "" Then
            If Trim(.txtKakuage.Text) <> "1" Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtKakuage, RED_CTL)
                Exit Function
            End If
            bXODC2 = True           '9/13 Yam
        End If
        ''���a�敪�`�F�b�N
        If Trim(.txtChokkei(0).Text) <> "" Or Trim(.txtChokkei(1).Text) <> "" Then
            If Trim(.txtChokkei(0).Text) > Trim(.txtChokkei(1).Text) Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtChokkei(0), RED_CTL)
                Call CtrlEnabled(.txtChokkei(1), RED_CTL)
                Exit Function
            End If
            bSIYO1 = True            '�d�lON
        End If
        ''�`���^�`�F�b�N
        .txtDendo.Text = UCase(.txtDendo.Text)
        If Trim(.txtDendo.Text) <> "" Then
            If Trim(.txtDendo.Text) <> "P" And Trim(.txtDendo.Text) <> "N" Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtDendo, RED_CTL)
                Exit Function
            End If
            bSIYO3 = True            '�d�lON
        End If
        ''�h�[�p���g�`�F�b�N
        If Trim(.txtDoba.Text) <> "" Then
            bSIYO3 = True            '�d�lON
        End If
        ''���ʃ`�F�b�N
        If Trim(.txtHoui.Text) <> "" Then
            bSIYO3 = True            '�d�lON
        End If
        ''��R���`�F�b�N
        If Trim(.txtTeikouritsu(0).Text) <> "" Then
            If Not IsNumeric(.txtTeikouritsu(0).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikouritsu(0), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '�d�lON
        End If
        If Trim(.txtTeikouritsu(1).Text) <> "" Then
            If Not IsNumeric(.txtTeikouritsu(1).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikouritsu(1), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '�d�lON
        End If
        If Trim(.txtTeikouritsu(0).Text) <> "" And Trim(.txtTeikouritsu(1).Text) <> "" Then
            If Val(Trim(.txtTeikouritsu(0).Text)) > Val(Trim(.txtTeikouritsu(1).Text)) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikouritsu(0), RED_CTL)
                Call CtrlEnabled(.txtTeikouritsu(1), RED_CTL)
                Exit Function
            End If
        End If
        ''��R���Q�ƃ`�F�b�N
        If Trim(.txtTeikoKbn.Text) <> "" And Trim(.txtTeikoKbn.Text) <> "1" And _
            (Trim(.txtTeikouritsu(0).Text) <> "" Or Trim(.txtTeikouritsu(1).Text) <> "") Then
            Call MsgOut(50)
            Call CtrlEnabled(.txtTeikoKbn, RED_CTL)
            Exit Function
        End If
        ''��R(�����W)�`�F�b�N
        If Trim(.txtTeikou(0).Text) <> "" Then
            If Not IsNumeric(.txtTeikou(0).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikou(0), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '�d�lON
        End If
        If Trim(.txtTeikou(1).Text) <> "" Then
            If Not IsNumeric(.txtTeikou(1).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikou(1), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '�d�lON
        End If
        If Trim(.txtTeikou(0).Text) <> "" And Trim(.txtTeikou(1).Text) <> "" Then
            If Val(Trim(.txtTeikou(0).Text)) > Val(Trim(.txtTeikou(1).Text)) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikou(0), RED_CTL)
                Call CtrlEnabled(.txtTeikou(1), RED_CTL)
                Exit Function
            End If
        End If
        ''�_�f�Z�x�`�F�b�N
        If Trim(.txtSanso(0).Text) <> "" Then
            If Not IsNumeric(.txtSanso(0).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtSanso(0), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '�d�lON
        End If
        If Trim(.txtSanso(1).Text) <> "" Then
            If Not IsNumeric(.txtSanso(1).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtSanso(1), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '�d�lON
        End If
        If Trim(.txtSanso(0).Text) <> "" And Trim(.txtSanso(1).Text) <> "" Then
            If Val(Trim(.txtSanso(0).Text)) > Val(Trim(.txtSanso(1).Text)) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtSanso(0), RED_CTL)
                Call CtrlEnabled(.txtSanso(1), RED_CTL)
                Exit Function
            End If
        End If
        ''�_�f�Z�x�Q�ƃ`�F�b�N
        If Trim(.txtSansoKbn.Text) <> "" And Trim(.txtSansoKbn.Text) <> "1" And _
            (Trim(.txtSanso(0).Text) <> "" Or Trim(.txtSanso(1).Text) <> "") Then
            Call MsgOut(50)
            Call CtrlEnabled(.txtSansoKbn, RED_CTL)
            Exit Function
        End If
        ''oi(�����W)�`�F�b�N
        If Trim(.txtOi(0).Text) <> "" Then
            If Not IsNumeric(.txtOi(0).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOi(0), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '�d�lON
        End If
        If Trim(.txtOi(1).Text) <> "" Then
            If Not IsNumeric(.txtOi(1).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOi(1), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '�d�lON
        End If
        If Trim(.txtOi(0).Text) <> "" And Trim(.txtOi(1).Text) <> "" Then
            If Val(Trim(.txtOi(0).Text)) > Val(Trim(.txtOi(1).Text)) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOi(0), RED_CTL)
                Call CtrlEnabled(.txtOi(1), RED_CTL)
                Exit Function
            End If
        End If
        ''ORG�`�F�b�N
        If Trim(.txtOrg(0).Text) <> "" Then
            If Not IsNumeric(.txtOrg(0).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOrg(0), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '�d�lON
        End If
        If Trim(.txtOrg(1).Text) <> "" Then
            If Not IsNumeric(.txtOrg(1).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOrg(1), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '�d�lON
        End If
        If Trim(.txtOrg(0).Text) <> "" And Trim(.txtOrg(1).Text) <> "" Then
            If Val(Trim(.txtOrg(0).Text)) > Val(Trim(.txtOrg(1).Text)) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOrg(0), RED_CTL)
                Call CtrlEnabled(.txtOrg(1), RED_CTL)
                Exit Function
            End If
        End If
    End With
    CheckF4Vbx5XX2 = True
End Function

' @(f)
' �@�\    : ���ԃ`�F�b�N�iVBX5XX2��p�j
' �Ԃ�l  :  OK - TRUE
'            NG -FALSE
' ������  : ctlControlS : �R���g���[��(�J�n��)
'           ctlControlE : �R���g���[��(�I����)
' �@�\����:

Private Function KikanCheckVbx5XX2(ctlControlS As Control, ctlControlE As Control) As Boolean
    'xxxxxxxxxxxxxxxxxxxxxxx
    '   mdlCommon.bas?
    'xxxxxxxxxxxxxxxxxxxxxxx
    Dim sDtS    As String       ''�W�v���ԊJ�n��
    Dim sDtE    As String       ''�W�v���ԏI����
    Dim sDtT    As String       ''�V�X�e�����t
    Dim sDtL    As String       ''�Y�����̌�����(�J�n��)
    Dim sDtLE   As String       ''�Y�����̌�����(�I����)
    Dim sWk     As String
    
    KikanCheckVbx5XX2 = False
    
    ''�W�v���Ԏ擾�ϊ�(6��[yymmdd]��8��[yyyymmdd])
    sDtS = DateChange(Trim(ctlControlS.Text))
    sDtE = DateChange(Trim(ctlControlE.Text))
    
    ''�J�n���E�I�����̐��l�`�F�b�N
    If Len(sDtS) = 8 Then
        If IsNumeric(Trim(sDtS)) = False Then
            Call CtrlEnabled(ctlControlS, RED_CTL)
            Call MsgOut(52, "", ERR_DISP)
            Exit Function
        End If
    ElseIf Len(sDtE) = 8 Then
        If IsNumeric(Trim(sDtE)) = False Then
            Call CtrlEnabled(ctlControlE, RED_CTL)
            Call MsgOut(52, "", ERR_DISP)
            Exit Function
        End If
    End If
    If Val(Mid(sDtS, 5, 2)) < 1 Or Val(Mid(sDtS, 5, 2)) > 12 Then
        Call CtrlEnabled(ctlControlS, RED_CTL)
        Call MsgOut(52, "", ERR_DISP)
        Exit Function
    ElseIf Val(Mid(sDtE, 5, 2)) < 1 Or Val(Mid(sDtE, 5, 2)) > 12 Then
        Call CtrlEnabled(ctlControlS, RED_CTL)
        Call MsgOut(52, "", ERR_DISP)
        Exit Function
    End If
    ''�������`�F�b�N�i�J�n�����I�����H�j
    If Val(sDtS) > Val(sDtE) Then
        Call CtrlEnabled(ctlControlS, RED_CTL)
        Call CtrlEnabled(ctlControlE, RED_CTL)
        Call MsgOut(53, "", ERR_DISP)
        Exit Function
    End If
    
    KikanCheckVbx5XX2 = True
End Function

' @(f)
' �@�\      : ��ʏ�����(Vbx5XX2)
' �Ԃ�l    : �Ȃ�
' ����      : bStatus�FTrue=�e�L�X�g�N���A
'                    �FFalse=�w�i�F�̂�
' �@�\����  : ���o��ʂ̏�����

Public Sub InitVbx5XX2(bStatus As Boolean)
    Dim i As Integer    ''���[�v�J�E���^

    With frmSub
        If bStatus Then
            ''���W�I�{�^�������ݒ�
            .optHinban(0).Value = True          ''�i�� ��v�s��v
            .optKisy(0).Value = True            ''�@�� ��v�s��v
            .optHikiageX(0).Value = True        ''������@ ��v�s��v
            .optPgid(0).Value = True            ''PGID ��v�s��v
            .optSeizoKbn(0).Value = True        ''���i�敪 ��v�s��v
            .optMokuteki(0).Value = True        ''�g�p�ړI ��v�s��v
            '.optMukaisaki(0).Value = True       ''���� ��v�s��v
        End If
        ''���̓t�B�[���h�̃N���A
        Call CtrlEnabled(.txtKisy, NORMAL_CTL, bStatus)            ''�@��
        Call CtrlEnabled(.txtPgid, NORMAL_CTL, bStatus)            ''PGID
        Call CtrlEnabled(.txtSeizoKbn, NORMAL_CTL, bStatus)        ''���i�敪
        Call CtrlEnabled(.txtKakuage, NORMAL_CTL, bStatus)         ''�i��敪
        Call CtrlEnabled(.txtDendo, NORMAL_CTL, bStatus)           ''�`���^
        Call CtrlEnabled(.txtDoba, NORMAL_CTL, bStatus)            ''�h�[�p���g
        Call CtrlEnabled(.txtHoui, NORMAL_CTL, bStatus)            ''����
        Call CtrlEnabled(.txtTeikoKbn, NORMAL_CTL, bStatus)        ''��R���i�����l�Q�Ɨ��j
        Call CtrlEnabled(.txtSansoKbn, NORMAL_CTL, bStatus)        ''�_�f�Z�x�i�����l�Q�Ɨ��j
        For i = 0 To 1
            Call CtrlEnabled(.txtSansoKbn, NORMAL_CTL, bStatus)        ''�_�f�Z�x
            Call CtrlEnabled(.txtHinban(i), NORMAL_CTL, bStatus)       ''�i��
            'Call CtrlEnabled(.txtMukaisaki(i), NORMAL_CTL, bStatus)    ''����
            'Call CtrlEnabled(.txtNoki(i), NORMAL_CTL, bStatus)         ''�[��
            Call CtrlEnabled(.txtGoki(i), NORMAL_CTL, bStatus)         ''���@
            Call CtrlEnabled(.txtChokkei(i), NORMAL_CTL, bStatus)      ''���a�敪
            Call CtrlEnabled(.txtTeikouritsu(i), NORMAL_CTL, bStatus)  ''��R��
            Call CtrlEnabled(.txtTeikou(i), NORMAL_CTL, bStatus)       ''�����W
            Call CtrlEnabled(.txtSanso(i), NORMAL_CTL, bStatus)        ''�_�f�Z�x
            Call CtrlEnabled(.txtOi(i), NORMAL_CTL, bStatus)           ''Oi�����W
            Call CtrlEnabled(.txtOrg(i), NORMAL_CTL, bStatus)          ''ORG
        Next i
        For i = 0 To 2
            Call CtrlEnabled(.txtHikiageX(i), NORMAL_CTL, bStatus)     ''������@
        Next i
        For i = 0 To 5
            Call CtrlEnabled(.txtMokuteki(i), NORMAL_CTL, bStatus)     ''�g�p�ړI
        Next i
    End With
End Sub

' @(f)
' �@�\      :   �@��E������@����{�ꕶ����ϊ�
' �Ԃ�l    :�@ TRUE �F����
'               FALSE�F�ُ�
' ����      :   iKbn�F�����敪  1:�@��E������@
'                               2:�@��E���@
'                               3:�@��E���@�E�i��
'               sCds�F(IN)�@��R�[�h���H�R�[�h[���i��]
'               sStr�F(OUT)�@�햼���H�R�[�h[���i��]
'
' �@�\����  :�@ �@�큕������@�^���@No[���i��]����Ǘ��R�[�h�̓��{��ɕ�����ϊ����A�߂��B
'
' ���l      :   '2000/08/15 ���� ���̏�����ǉ������B

Public Function ChgKisyuStr(ByVal iKbn As Integer, ByVal sCds As String, ByRef sStr As String) As Boolean
    Static bReadad As Boolean
    Static Kisyu() As CDNAMEDAT
    Static Pumeth() As CDNAMEDAT
    
    Dim sKisyu As String
    Dim sCd1 As String
    Dim sCd2 As String
    Dim sSQL As String
    Dim objOraDyn As OraDynaset             ''�_�C�i�Z�b�g
    Dim iIdx As Integer
    Dim wk_Hinb As String
    
'    Debug.Print "�R�[�h�F" & sCds
    ChgKisyuStr = False
    
    ''�@��E������@���܂��擾���ĂȂ����
    If Not bReadad Then
        ''�@�햼�擾����SQL���쐬
        sSQL = "SELECT NVL(  codea9,   ' '), "   ''�ʃR�[�h
        sSQL = sSQL & "NVL(  namesja9, ' ')  "   ''�R�[�h���i���{��Z�k�j
        sSQL = sSQL & "FROM  koda9           "
        sSQL = sSQL & "WHERE shuca9 = '44'   "
        sSQL = sSQL & "  AND sysca9 = 'K'    "
        ''�_�C�i�Z�b�g�쐬�i�������s�j
        If DynSet(objOraDyn, sSQL) = False Then
            ''�_�C�i�Z�b�g�쐬���s
            Call MsgOut(100, "", ERR_DISP_LOG, "KODA9")
            ''�������~
            Exit Function
        End If
        ''�Ĕz�u
        ReDim Kisyu(objOraDyn.RecordCount)
        ''�擾�m�F
        iIdx = 0
        Do While Not objOraDyn.EOF
            With Kisyu(iIdx)
                .Cd = Trim(objOraDyn(0))    ''�@��R�[�h
                .Nm = Trim(objOraDyn(1))    ''�@�햼�擾
            End With
            iIdx = iIdx + 1
            objOraDyn.MoveNext
        Loop
        
        ''������@�擾����SQL���쐬
        sSQL = "SELECT NVL(  codea9,  ' '), "            ''�R�[�h���i���{��j
        sSQL = sSQL & "NVL(  nameja9, ' ')  "
        sSQL = sSQL & "FROM  koda9          "
        sSQL = sSQL & "WHERE shuca9 = '51'  "
        sSQL = sSQL & "  AND sysca9 = 'X'   "
        ''�_�C�i�Z�b�g�쐬�i�������s�j
        If DynSet(objOraDyn, sSQL) = False Then
            ''�_�C�i�Z�b�g�쐬���s
            Call MsgOut(100, "", ERR_DISP_LOG, "KODA9")
            ''�������~
            Exit Function
        End If
        ReDim Pumeth(objOraDyn.RecordCount)
        ''�擾�m�F
        iIdx = 0
        Do While Not objOraDyn.EOF
            With Pumeth(iIdx)
                .Cd = Trim(objOraDyn(0))    ''�@��R�[�h
                .Nm = Trim(objOraDyn(1))    ''�@�햼�擾
            End With
            iIdx = iIdx + 1
            objOraDyn.MoveNext
        Loop
        
        ''�Ǎ���
        bReadad = True
    End If
    
    ''�R�[�h�؂�o��
    sKisyu = Left(sCds, 2)      ''�@��@�؂�o��
    Select Case iKbn
    Case 1                      ''�@��ʑI������
        sCd1 = Mid(sCds, 3, 1)  ''������@ �؂�o��
        sCd2 = ""
    Case 2                      ''���@�ʑI������
        sCd1 = Mid(sCds, 3)     ''���@ �؂�o��
        sCd2 = ""
    Case 3                      ''���@�i�ԕʑI������
        sCd1 = Mid(sCds, 3, 3)  ''���@ �؂�o��
        sCd2 = Mid(sCds, 6)     ''�i�� �؂�o��
    Case 5                      ''�@��ʑI������
        sCd1 = Mid(sCds, 4, 1)  ''������@ �؂�o��
        sCd2 = ""
    Case 6                      ''���@�ʑI������
        sCd1 = Mid(sCds, 4)     ''���@ �؂�o��
        sCd2 = ""
    End Select
    
    ''�@�팟���E���O�擾
    For iIdx = 0 To UBound(Kisyu)
        If Kisyu(iIdx).Cd = sKisyu Then sKisyu = Kisyu(iIdx).Nm: Exit For
    Next
    If (iKbn And 3) = 1 Then    ''������@�Ȃ�
        ''������@�����E���O�擾
        For iIdx = 0 To UBound(Pumeth)
            If Pumeth(iIdx).Cd = sCd1 Then sCd1 = Pumeth(iIdx).Nm: Exit For
        Next
    End If
    
    ''�߂�
    'sStr = sKisyu & " " & sCd1 & " " & Format(sCd2, "!@@@-@@@@-@@@@") & vbTab
    If GetHinbanHensyu(Trim(sCd2), 1, wk_Hinb) = True Then
        sStr = sKisyu & " " & sCd1 & " " & wk_Hinb & Chr(9) ' vbTab
    End If
    ChgKisyuStr = True
End Function

' @(f)
' �@�\      :   ���i�R�[�h���疼�̂��擾
' �Ԃ�l    :�@ TRUE �F����
'               FALSE�F�ُ�
' ����      :   ���i�R�[�h
' �@�\����  :�@ ���i�R�[�h���疼�̂��擾

Public Function GetGreadName(ByVal sCds As String, ByRef sStr As String) As Boolean
    Dim sName As String
    Dim sCd1 As String
    Dim sSQL As String
    Dim objOraDyn As OraDynaset             ''�_�C�i�Z�b�g
    
    GetGreadName = False
    
    ''���i�R�[�h���擾����SQL���쐬
    sSQL = "SELECT NVL(namjls9,' ') "   ''����
    sSQL = sSQL & "FROM  sods9           "
    sSQL = sSQL & "WHERE nams9 = 'COMPRDGRS1'   "
    sSQL = sSQL & "  AND clss9 = '01'    "
    sSQL = sSQL & "  AND vals9 = '" & sCds & "'    "
    ''�_�C�i�Z�b�g�쐬�i�������s�j
    If DynSet(objOraDyn, sSQL) = False Then
        ''�_�C�i�Z�b�g�쐬���s
        Call MsgOut(100, "", ERR_DISP_LOG, "sods9")
        ''�������~
        Exit Function
    End If
    If objOraDyn.EOF = True Then   '4/19 Yam �ǉ�
        sStr = " " & Chr(9)
        GetGreadName = True
        Exit Function
    End If

    sName = objOraDyn(0)
    
    sStr = sCds & " " & sName & Chr(9)
    
    GetGreadName = True
End Function

' @(f)
' �@�\      :   �@��R�[�h�ϊ�
' �Ԃ�l    :�@ TRUE �F����
'               FALSE�F�ُ�
' �@�\����  :�@ �@��Ǘ��R�[�h����ʃR�[�h�ɕϊ����A�߂��B
' ���l      :   '2000/11/17 Yam

Public Function GetkisyNo(ByVal bkisy As String, ByRef akisy As String) As Boolean
    Dim sSQL As String
    Dim objOraDyn As OraDynaset             ''�_�C�i�Z�b�g
    
    GetkisyNo = False
        ''�@�햼�擾����SQL���쐬
        sSQL = "SELECT NVL(  codea9,   ' ') "   ''�ʃR�[�h
        sSQL = sSQL & "FROM  koda9           "
        sSQL = sSQL & "WHERE shuca9 = '44'   "
        sSQL = sSQL & "  AND sysca9 = 'K'    "
        sSQL = sSQL & "  AND kcodea9 =  '" & Trim(bkisy) & "'"
        Debug.Print sSQL
        ''�_�C�i�Z�b�g�쐬�i�������s�j
        If DynSet(objOraDyn, sSQL) = False Then
            ''�_�C�i�Z�b�g�쐬���s
            Call MsgOut(100, "", ERR_DISP_LOG, "KODA9")
            ''�������~
            Exit Function
        End If
        If objOraDyn.EOF = True Then
            Exit Function
        End If
         
        akisy = objOraDyn(0)    ''�ʃR�[�h
                
    GetkisyNo = True
End Function
